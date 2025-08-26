Option Explicit

' =============================================================================
' DATE ANALYSIS MACRO - SIMPLIFIED VERSION
' =============================================================================
' Purpose: Analyze task completion times by person and month
' Input: Data sheet with columns A=keyword, B=name, C=start date, D=end date
' Output: Monthly summary tables and chart showing on-time percentages
' =============================================================================

' Configuration - Change these values as needed
Private Const START_DATE As String = "01/01/2025"        ' Analysis start date
Private Const END_DATE As String = "08/25/2025"          ' Analysis end date
Private Const NAME_COLUMN As String = "B"                ' Column containing person names
Private Const START_DATE_COLUMN As String = "C"          ' Column containing task start dates
Private Const END_DATE_COLUMN As String = "D"            ' Column containing task end dates
Private Const KEYWORD_COLUMN As String = "A"             ' Column containing keywords for filtering
Private Const KEYWORD_TO_FILTER As String = "tech"       ' Filter by this keyword (leave empty for all)
Private Const DAYS_THRESHOLD_1 As Long = 7               ' "On time" threshold (days)
Private Const DAYS_THRESHOLD_2 As Long = 14              ' "Acceptable" threshold (days)

Sub AnalyzeDatesWithMonthlyOutputsAndStackedHistogram()
    ' =============================================================================
    ' MAIN VARIABLES - All variables declared up front for clarity
    ' =============================================================================
    
    ' Worksheet objects
    Dim wsData As Worksheet, wsOutput As Worksheet
    
    ' Date variables
    Dim startDate As Date, endDate As Date
    Dim cellStartDate As Variant, cellEndDate As Variant
    
    ' Loop counters and data processing
    Dim lastRow As Long, currentRow As Long, outputRow As Long
    Dim dateDiff As Long, invalidRows As Long
    Dim nameKey As String, monthKey As String
    
    ' Data storage dictionaries - think of these as smart tables that auto-expand
    Dim dictMonthly7Days As Object      ' Stores count of tasks <= 7 days by person/month
    Dim dictMonthly14Days As Object     ' Stores count of tasks <= 14 days by person/month
    Dim dictMonthly14PlusDays As Object ' Stores count of tasks > 14 days by person/month
    Dim allNames As Object              ' List of all unique person names
    Dim allMonths As Object             ' List of all unique months
    
    ' Arrays for sorted data
    Dim sortedMonths() As String, nameKeys() As String
    Dim idx As Long, nameIdx As Long
    
    ' Calculation variables
    Dim count7 As Long, count14 As Long, count14Plus As Long, totalCount As Long
    Dim percentageOnTime As Double
    
    ' Chart variables
    Dim co As ChartObject, chartDataRange As Range
    Dim seriesIndex As Long, colorIndex As Long
    
    ' =============================================================================
    ' SETUP AND INITIALIZATION
    ' =============================================================================
    
    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    
    ' Convert date strings to actual dates
    startDate = CDate(START_DATE)
    endDate = CDate(END_DATE)
    
    ' Get the source data worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    
    ' Create dictionaries to store our data (like expandable tables)
    Set dictMonthly7Days = CreateObject("Scripting.Dictionary")
    Set dictMonthly14Days = CreateObject("Scripting.Dictionary")
    Set dictMonthly14PlusDays = CreateObject("Scripting.Dictionary")
    Set allNames = CreateObject("Scripting.Dictionary")
    Set allMonths = CreateObject("Scripting.Dictionary")
    
    ' Find the last row with data in the name column
    lastRow = wsData.Cells(wsData.Rows.Count, NAME_COLUMN).End(xlUp).row
    invalidRows = 0
    
    ' =============================================================================
    ' DATA PROCESSING LOOP - Read each row and categorize the data
    ' =============================================================================
    
    For currentRow = 2 To lastRow ' Start at row 2 (skip headers)
        
        ' Get data from current row - using configurable column constants
        Dim cellName As Variant: cellName = wsData.Cells(currentRow, NAME_COLUMN).Value           ' Person name
        cellStartDate = wsData.Cells(currentRow, START_DATE_COLUMN).Value                         ' Task start date
        cellEndDate = wsData.Cells(currentRow, END_DATE_COLUMN).Value                             ' Task end date
        Dim cellKeyword As Variant: cellKeyword = wsData.Cells(currentRow, KEYWORD_COLUMN).Value  ' Keyword filter
        
        ' Skip this row if data is invalid
        If IsEmpty(cellName) Or Not IsDate(cellStartDate) Or Not IsDate(cellEndDate) Then
            invalidRows = invalidRows + 1
            GoTo NextRow
        End If
        
        ' Apply keyword filter (skip if doesn't match)
        If KEYWORD_TO_FILTER <> "" Then
            If InStr(1, LCase(CStr(cellKeyword)), LCase(KEYWORD_TO_FILTER)) = 0 Then
                invalidRows = invalidRows + 1
                GoTo NextRow
            End If
        End If
        
        ' Only process data within our date range
        If CDate(cellStartDate) >= startDate And CDate(cellStartDate) <= endDate Then
            
            ' Calculate how many days the task took
            dateDiff = CDate(cellEndDate) - CDate(cellStartDate)
            
            ' Create keys for organizing data
            nameKey = CStr(cellName)                                    ' Person's name
            monthKey = Format(CDate(cellStartDate), "YYYY-MM")          ' Month in YYYY-MM format
            
            ' Remember this person and month for later
            allNames(nameKey) = True
            allMonths(monthKey) = True
            
            ' Create nested dictionaries if they don't exist (Person -> Month -> Count)
            If Not dictMonthly7Days.Exists(nameKey) Then Set dictMonthly7Days(nameKey) = CreateObject("Scripting.Dictionary")
            If Not dictMonthly14Days.Exists(nameKey) Then Set dictMonthly14Days(nameKey) = CreateObject("Scripting.Dictionary")
            If Not dictMonthly14PlusDays.Exists(nameKey) Then Set dictMonthly14PlusDays(nameKey) = CreateObject("Scripting.Dictionary")
            
            ' Categorize the task based on how long it took and increment counters
            If dateDiff <= DAYS_THRESHOLD_1 Then
                ' Task completed within 7 days - "On time"
                If Not dictMonthly7Days(nameKey).Exists(monthKey) Then dictMonthly7Days(nameKey)(monthKey) = 0
                dictMonthly7Days(nameKey)(monthKey) = dictMonthly7Days(nameKey)(monthKey) + 1
                
            ElseIf dateDiff <= DAYS_THRESHOLD_2 Then
                ' Task completed within 8-14 days - "Acceptable"
                If Not dictMonthly14Days(nameKey).Exists(monthKey) Then dictMonthly14Days(nameKey)(monthKey) = 0
                dictMonthly14Days(nameKey)(monthKey) = dictMonthly14Days(nameKey)(monthKey) + 1
                
            Else
                ' Task took more than 14 days - "Late"
                If Not dictMonthly14PlusDays(nameKey).Exists(monthKey) Then dictMonthly14PlusDays(nameKey)(monthKey) = 0
                dictMonthly14PlusDays(nameKey)(monthKey) = dictMonthly14PlusDays(nameKey)(monthKey) + 1
            End If
        End If
NextRow:
    Next currentRow
    
    ' =============================================================================
    ' CREATE OUTPUT WORKSHEET
    ' =============================================================================
    
    ' Delete old results sheet if it exists, then create new one
    On Error Resume Next
    ThisWorkbook.Sheets("MonthlySummary").Delete
    On Error GoTo 0
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Data"))
    wsOutput.name = "MonthlySummary"
    
    ' =============================================================================
    ' CREATE SORTED ARRAYS - Convert dictionary keys to sorted arrays for consistent output
    ' =============================================================================
    
    ' Create sorted array of months
    If allMonths.Count > 0 Then
        ReDim sortedMonths(0 To allMonths.Count - 1)
        idx = 0
        Dim monthVar As Variant
        For Each monthVar In allMonths.Keys
            sortedMonths(idx) = CStr(monthVar)
            idx = idx + 1
        Next
        SortStringArray sortedMonths ' Sort chronologically
    End If
    
    ' Create sorted array of names
    If allNames.Count > 0 Then
        ReDim nameKeys(0 To allNames.Count - 1)
        idx = 0
        Dim nameVar As Variant
        For Each nameVar In allNames.Keys
            nameKeys(idx) = CStr(nameVar)
            idx = idx + 1
        Next
        SortStringArray nameKeys ' Sort alphabetically
    End If
    
    ' =============================================================================
    ' WRITE MONTHLY SUMMARY TABLES - The main output tables
    ' =============================================================================
    
    ' Write main table headers
    outputRow = 1
    wsOutput.Cells(outputRow, 1).Value = "Name"
    wsOutput.Cells(outputRow, 2).Value = "Period"
    wsOutput.Cells(outputRow, 3).Value = "<= 7 Days"
    wsOutput.Cells(outputRow, 4).Value = "<= 14 Days"
    wsOutput.Cells(outputRow, 5).Value = "> 14 Days"
    wsOutput.Cells(outputRow, 6).Value = "Percentage On Time"
    
    ' Format headers
    With wsOutput.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    outputRow = outputRow + 1
    
    ' Loop through each month and create a section
    For idx = 0 To UBound(sortedMonths)
        monthKey = sortedMonths(idx)
        
        ' Month section header
        wsOutput.Cells(outputRow, 1).Value = "Month: " & monthKey
        With wsOutput.Range(wsOutput.Cells(outputRow, 1), wsOutput.Cells(outputRow, 6))
            .Merge
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(180, 180, 180)
            .HorizontalAlignment = xlCenter
        End With
        outputRow = outputRow + 1
        
        ' Track monthly totals across all people
        Dim monthTotal7 As Long, monthTotal14 As Long, monthTotal14Plus As Long
        monthTotal7 = 0: monthTotal14 = 0: monthTotal14Plus = 0
        
        ' Loop through each person for this month
        For nameIdx = 0 To UBound(nameKeys)
            nameKey = nameKeys(nameIdx)
            
            ' Get counts for this person in this month
            count7 = 0: count14 = 0: count14Plus = 0
            If dictMonthly7Days.Exists(nameKey) And dictMonthly7Days(nameKey).Exists(monthKey) Then count7 = dictMonthly7Days(nameKey)(monthKey)
            If dictMonthly14Days.Exists(nameKey) And dictMonthly14Days(nameKey).Exists(monthKey) Then count14 = dictMonthly14Days(nameKey)(monthKey)
            If dictMonthly14PlusDays.Exists(nameKey) And dictMonthly14PlusDays(nameKey).Exists(monthKey) Then count14Plus = dictMonthly14PlusDays(nameKey)(monthKey)
            
            totalCount = count7 + count14 + count14Plus
            
            ' Only write row if person had tasks this month
            If totalCount > 0 Then
                wsOutput.Cells(outputRow, 1).Value = nameKey
                wsOutput.Cells(outputRow, 2).Value = monthKey
                wsOutput.Cells(outputRow, 3).Value = count7
                wsOutput.Cells(outputRow, 4).Value = count14
                wsOutput.Cells(outputRow, 5).Value = count14Plus
                
                ' Calculate on-time percentage (7 days + 14 days = acceptable)
                percentageOnTime = (count7 + count14) / totalCount
                wsOutput.Cells(outputRow, 6).Value = percentageOnTime
                wsOutput.Cells(outputRow, 6).NumberFormat = "0.00%"
                
                ' Add to monthly totals
                monthTotal7 = monthTotal7 + count7
                monthTotal14 = monthTotal14 + count14
                monthTotal14Plus = monthTotal14Plus + count14Plus
                
                outputRow = outputRow + 1
            End If
        Next nameIdx
        
        ' Write monthly total row
        wsOutput.Cells(outputRow, 1).Value = "TOTAL"
        wsOutput.Cells(outputRow, 2).Value = monthKey
        wsOutput.Cells(outputRow, 3).Value = monthTotal7
        wsOutput.Cells(outputRow, 4).Value = monthTotal14
        wsOutput.Cells(outputRow, 5).Value = monthTotal14Plus
        
        ' Calculate monthly total percentage
        Dim monthGrandTotal As Long
        monthGrandTotal = monthTotal7 + monthTotal14 + monthTotal14Plus
        If monthGrandTotal > 0 Then
            wsOutput.Cells(outputRow, 6).Value = (monthTotal7 + monthTotal14) / monthGrandTotal
        Else
            wsOutput.Cells(outputRow, 6).Value = 0
        End If
        wsOutput.Cells(outputRow, 6).NumberFormat = "0.00%"
        
        ' Format total row
        With wsOutput.Range("A" & outputRow & ":F" & outputRow)
            .Font.Bold = True
            .Interior.Color = RGB(240, 240, 240)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
        
        outputRow = outputRow + 2 ' Leave space between months
    Next idx
    
    ' =============================================================================
    ' CREATE CHART DATA TABLE - Restructure data for charting
    ' =============================================================================
    
    ' Add space and create chart data section
    Dim chartTableStartRow As Long
    chartTableStartRow = outputRow + 2
    
    ' Chart data section header
    wsOutput.Cells(chartTableStartRow, 1).Value = "Chart Data Table:"
    With wsOutput.Cells(chartTableStartRow, 1)
        .Font.Bold = True
        .Font.Size = 12
    End With
    chartTableStartRow = chartTableStartRow + 1
    
    ' Chart data table headers: Month, Person1, Person2, etc.
    wsOutput.Cells(chartTableStartRow, 1).Value = "Month"
    For nameIdx = 0 To UBound(nameKeys)
        wsOutput.Cells(chartTableStartRow, nameIdx + 2).Value = nameKeys(nameIdx)
    Next nameIdx
    
    ' Format chart table headers
    With wsOutput.Range(wsOutput.Cells(chartTableStartRow, 1), wsOutput.Cells(chartTableStartRow, UBound(nameKeys) + 2))
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Fill chart data table - each row is a month, each column is a person
    Dim dataStartRow As Long, dataRowCount As Long
    dataStartRow = chartTableStartRow + 1
    dataRowCount = 0
    
    For idx = 0 To UBound(sortedMonths)
        monthKey = sortedMonths(idx)
        dataRowCount = dataRowCount + 1
        Dim currentDataRow As Long
        currentDataRow = dataStartRow + dataRowCount - 1
        
        ' First column: Month name
        wsOutput.Cells(currentDataRow, 1).Value = monthKey
        
        ' Remaining columns: Each person's on-time percentage for this month
        For nameIdx = 0 To UBound(nameKeys)
            nameKey = nameKeys(nameIdx)
            
            ' Get counts for this person/month combination
            count7 = 0: count14 = 0: count14Plus = 0
            If dictMonthly7Days.Exists(nameKey) And dictMonthly7Days(nameKey).Exists(monthKey) Then count7 = dictMonthly7Days(nameKey)(monthKey)
            If dictMonthly14Days.Exists(nameKey) And dictMonthly14Days(nameKey).Exists(monthKey) Then count14 = dictMonthly14Days(nameKey)(monthKey)
            If dictMonthly14PlusDays.Exists(nameKey) And dictMonthly14PlusDays(nameKey).Exists(monthKey) Then count14Plus = dictMonthly14PlusDays(nameKey)(monthKey)
            
            totalCount = count7 + count14 + count14Plus
            If totalCount > 0 Then
                ' Calculate on-time percentage
                percentageOnTime = CDbl((count7 + count14)) / CDbl(totalCount)
                wsOutput.Cells(currentDataRow, nameIdx + 2).Value = percentageOnTime
                wsOutput.Cells(currentDataRow, nameIdx + 2).NumberFormat = "0.00%"
            Else
                ' No data for this person/month - show 0%
                wsOutput.Cells(currentDataRow, nameIdx + 2).Value = 0
                wsOutput.Cells(currentDataRow, nameIdx + 2).NumberFormat = "0.00%"
            End If
        Next nameIdx
    Next idx
    
    ' Add borders around chart data table
    With wsOutput.Range(wsOutput.Cells(chartTableStartRow, 1), wsOutput.Cells(dataStartRow + dataRowCount - 1, UBound(nameKeys) + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' =============================================================================
    ' CREATE CHART - Visual representation of the data
    ' =============================================================================
    
    ' Define the data range for the chart (the table we just created)
    Set chartDataRange = wsOutput.Range(wsOutput.Cells(chartTableStartRow, 1), wsOutput.Cells(dataStartRow + dataRowCount - 1, UBound(nameKeys) + 2))
    
    ' Create chart below the data table
    Dim chartTop As Long
    chartTop = wsOutput.Cells(dataStartRow + dataRowCount + 2, 1).Top
    Set co = wsOutput.ChartObjects.Add(Left:=50, Top:=chartTop, Width:=600, Height:=300)
    
    ' Configure the chart
    With co.Chart
        .ChartType = xlColumnClustered                      ' Clustered column chart (grouped bars)
        .SetSourceData Source:=chartDataRange, PlotBy:=xlColumns ' Each column (person) becomes a data series
        
        ' Chart title
        .HasTitle = True
        .ChartTitle.Text = "On-Time Percentage by Month and Name"
        
        ' X-axis (shows months)
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Month"
            .TickLabels.Orientation = 45 ' Rotate labels for better readability
        End With
        
        ' Y-axis (shows percentages)
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Percentage On Time"
            .TickLabels.NumberFormat = "0%" ' Show as percentages
            .MinimumScale = 0
            .MaximumScale = 1
        End With
        
        ' Legend (shows which color = which person)
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        
        ' Color each data series (person) with different colors
        If .SeriesCollection.Count > 0 Then
            For seriesIndex = 1 To .SeriesCollection.Count
                With .SeriesCollection(seriesIndex)
                    ' Set the series name to the person's name
                    .name = nameKeys(seriesIndex - 1)
                    
                    ' Apply a unique color (cycle through 15 colors, then generate more)
                    colorIndex = ((seriesIndex - 1) Mod 15) + 1
                    Select Case colorIndex
                        Case 1: .Interior.Color = RGB(68, 114, 196)    ' Blue
                        Case 2: .Interior.Color = RGB(237, 125, 49)    ' Orange
                        Case 3: .Interior.Color = RGB(112, 173, 71)    ' Green
                        Case 4: .Interior.Color = RGB(255, 192, 0)     ' Yellow
                        Case 5: .Interior.Color = RGB(91, 155, 213)    ' Light Blue
                        Case 6: .Interior.Color = RGB(165, 165, 165)   ' Gray
                        Case 7: .Interior.Color = RGB(158, 72, 14)     ' Brown
                        Case 8: .Interior.Color = RGB(99, 99, 99)      ' Dark Gray
                        Case 9: .Interior.Color = RGB(153, 115, 0)     ' Olive
                        Case 10: .Interior.Color = RGB(67, 104, 43)    ' Dark Green
                        Case 11: .Interior.Color = RGB(196, 89, 17)    ' Dark Orange
                        Case 12: .Interior.Color = RGB(142, 169, 219)  ' Periwinkle
                        Case 13: .Interior.Color = RGB(255, 105, 180)  ' Hot Pink
                        Case 14: .Interior.Color = RGB(32, 178, 170)   ' Light Sea Green
                        Case 15: .Interior.Color = RGB(128, 0, 128)    ' Purple
                        Case Else
                            ' For more than 15 people, generate colors mathematically
                            .Interior.Color = RGB((seriesIndex * 50) Mod 256, (seriesIndex * 80) Mod 256, (seriesIndex * 110) Mod 256)
                    End Select
                End With
            Next seriesIndex
        End If
    End With
    
    ' =============================================================================
    ' CLEANUP AND COMPLETION
    ' =============================================================================
    
    ' Auto-fit all columns for better display
    wsOutput.Columns("A:Z").AutoFit
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    ' Show completion message
    MsgBox "Analysis complete!" & vbCrLf & _
           "Date Range: " & Format(startDate, "MM/DD/YYYY") & " to " & Format(endDate, "MM/DD/YYYY") & vbCrLf & _
           "Rows Processed: " & (lastRow - 1 - invalidRows) & vbCrLf & _
           "Results are in the 'MonthlySummary' sheet.", vbInformation
End Sub

' =============================================================================
' HELPER FUNCTIONS
' =============================================================================

' Simple bubble sort to sort string arrays alphabetically/chronologically
Sub SortStringArray(arr() As String)
    Dim i As Long, j As Long, temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

