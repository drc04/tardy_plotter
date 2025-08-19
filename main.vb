Function WeekOfMonth(ByVal d As Date) As Long
    ' Calculates the week number within a given month.
    Dim firstDay As Date
    firstDay = DateSerial(Year(d), Month(d), 1)
    
    ' Calculate the week number based on the first day of the month
    WeekOfMonth = Application.Floor((Day(d) - 1 + Weekday(firstDay, vbMonday) - 1) / 7) + 1
End Function

Sub AnalyzeDatesWithFinalMetricsAndChart()
    ' Define variables
    Dim wsData As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, currentRow As Long
    Dim dateDiff As Long
    Dim startDate As Date, endDate As Date
    Dim nameKey As String, monthKey As String, weekNumber As Long
    Dim chartSheet As Worksheet
    
    ' Dictionaries for monthly and weekly totals
    Dim dictMonthly7Days As Object, dictMonthly14Days As Object, dictMonthly14PlusDays As Object
    Dim dictWeekly7Days As Object, dictWeekly14Days As Object, dictWeekly14PlusDays As Object
    Dim dictTotalCounts As Object
    
    ' Arrays to store monthly data for charting
    Dim namesArray() As String
    Dim datesArray() As String
    Dim percentagesArray() As Double
    Dim nameIndex As Object, dateIndex As Object
    Dim totalNames As Long, totalMonths As Long
    
    On Error GoTo SheetError
    Set wsData = ThisWorkbook.Sheets("Data") ' 🚨 Change "Data" to your sheet name
    On Error GoTo 0

    On Error GoTo Canceled
    startDate = Application.InputBox("Enter the start date (MM/DD/YYYY):", "Start Date", Type:=1)
    endDate = Application.InputBox("Enter the end date (MM/DD/YYYY):", "End Date", Type:=1)
    If startDate = 0 Or endDate = 0 Then Exit Sub
    On Error GoTo 0

    Set dictMonthly7Days = CreateObject("Scripting.Dictionary")
    Set dictMonthly14Days = CreateObject("Scripting.Dictionary")
    Set dictMonthly14PlusDays = CreateObject("Scripting.Dictionary")
    Set dictWeekly7Days = CreateObject("Scripting.Dictionary")
    Set dictWeekly14Days = CreateObject("Scripting.Dictionary")
    Set dictWeekly14PlusDays = CreateObject("Scripting.Dictionary")
    Set dictTotalCounts = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    For currentRow = 2 To lastRow
        If wsData.Cells(currentRow, "B").Value >= startDate And wsData.Cells(currentRow, "B").Value <= endDate Then
            dateDiff = wsData.Cells(currentRow, "C").Value - wsData.Cells(currentRow, "B").Value
            nameKey = wsData.Cells(currentRow, "A").Value
            monthKey = Format(wsData.Cells(currentRow, "B").Value, "YYYY-MM")
            weekNumber = WeekOfMonth(wsData.Cells(currentRow, "B").Value)
            
            If Not dictTotalCounts.Exists(nameKey) Then
                Set dictTotalCounts(nameKey) = CreateObject("Scripting.Dictionary")
                dictTotalCounts(nameKey).Add "Less7", 0
                dictTotalCounts(nameKey).Add "Less14", 0
                dictTotalCounts(nameKey).Add "Greater14", 0
            End If
            
            If dateDiff <= 7 Then
                If Not dictMonthly7Days.Exists(nameKey) Then Set dictMonthly7Days(nameKey) = CreateObject("Scripting.Dictionary")
                dictMonthly7Days(nameKey)(monthKey) = dictMonthly7Days(nameKey)(monthKey) + 1
                If Not dictWeekly7Days.Exists(nameKey) Then Set dictWeekly7Days(nameKey) = CreateObject("Scripting.Dictionary")
                dictWeekly7Days(nameKey)(monthKey & "_Week_" & weekNumber) = dictWeekly7Days(nameKey)(monthKey & "_Week_" & weekNumber) + 1
                dictTotalCounts(nameKey)("Less7") = dictTotalCounts(nameKey)("Less7") + 1
            ElseIf dateDiff <= 14 Then
                If Not dictMonthly14Days.Exists(nameKey) Then Set dictMonthly14Days(nameKey) = CreateObject("Scripting.Dictionary")
                dictMonthly14Days(nameKey)(monthKey) = dictMonthly14Days(nameKey)(monthKey) + 1
                If Not dictWeekly14Days.Exists(nameKey) Then Set dictWeekly14Days(nameKey) = CreateObject("Scripting.Dictionary")
                dictWeekly14Days(nameKey)(monthKey & "_Week_" & weekNumber) = dictWeekly14Days(nameKey)(monthKey & "_Week_" & weekNumber) + 1
                dictTotalCounts(nameKey)("Less14") = dictTotalCounts(nameKey)("Less14") + 1
            Else
                If Not dictMonthly14PlusDays.Exists(nameKey) Then Set dictMonthly14PlusDays(nameKey) = CreateObject("Scripting.Dictionary")
                dictMonthly14PlusDays(nameKey)(monthKey) = dictMonthly14PlusDays(nameKey)(monthKey) + 1
                If Not dictWeekly14PlusDays.Exists(nameKey) Then Set dictWeekly14PlusDays(nameKey) = CreateObject("Scripting.Dictionary")
                dictWeekly14PlusDays(nameKey)(monthKey & "_Week_" & weekNumber) = dictWeekly14PlusDays(nameKey)(monthKey & "_Week_" & weekNumber) + 1
                dictTotalCounts(nameKey)("Greater14") = dictTotalCounts(nameKey)("Greater14") + 1
            End If
        End If
    Next currentRow

    Application.ScreenUpdating = False
    
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsData)
    wsOutput.Name = "DateRangeSummary"

    ' --- Output Monthly/Weekly Data ---
    Dim allNames As Object, allMonths As Object, allWeeks As Object
    Set allNames = CreateObject("Scripting.Dictionary")
    
    For Each nameKey In dictMonthly7Days.Keys: allNames(nameKey) = True: Next
    For Each nameKey In dictMonthly14Days.Keys: allNames(nameKey) = True: Next
    For Each nameKey In dictMonthly14PlusDays.Keys: allNames(nameKey) = True: Next
    For Each nameKey In dictWeekly7Days.Keys: allNames(nameKey) = True: Next
    For Each nameKey In dictWeekly14Days.Keys: allNames(nameKey) = True: Next
    For Each nameKey In dictWeekly14PlusDays.Keys: allNames(nameKey) = True: Next
    
    Dim row_num As Long
    row_num = 1
    
    wsOutput.Cells(row_num, 1).Value = "Name"
    wsOutput.Cells(row_num, 2).Value = "Period"
    wsOutput.Cells(row_num, 3).Value = "<= 7 Days"
    wsOutput.Cells(row_num, 4).Value = "<= 14 Days"
    wsOutput.Cells(row_num, 5).Value = "> 14 Days"
    wsOutput.Cells(row_num, 6).Value = "Percentage Late"
    row_num = row_num + 1

    For Each nameKey In allNames.Keys
        Set allMonths = CreateObject("Scripting.Dictionary")
        If dictMonthly7Days.Exists(nameKey) Then For Each subKey In dictMonthly7Days(nameKey).Keys: allMonths(subKey) = True: Next
        If dictMonthly14Days.Exists(nameKey) Then For Each subKey In dictMonthly14Days(nameKey).Keys: allMonths(subKey) = True: Next
        If dictMonthly14PlusDays.Exists(nameKey) Then For Each subKey In dictMonthly14PlusDays(nameKey).Keys: allMonths(subKey) = True: Next
        For Each monthKey In allMonths.Keys
            Dim count7 As Long, count14 As Long, count14Plus As Long, totalCount As Long, percentageLate As Double
            count7 = IIf(dictMonthly7Days.Exists(nameKey) And dictMonthly7Days(nameKey).Exists(monthKey), dictMonthly7Days(nameKey)(monthKey), 0)
            count14 = IIf(dictMonthly14Days.Exists(nameKey) And dictMonthly14Days(nameKey).Exists(monthKey), dictMonthly14Days(nameKey)(monthKey), 0)
            count14Plus = IIf(dictMonthly14PlusDays.Exists(nameKey) And dictMonthly14PlusDays(nameKey).Exists(monthKey), dictMonthly14PlusDays(nameKey)(monthKey), 0)
            totalCount = count7 + count14 + count14Plus
            If totalCount > 0 Then percentageLate = count14Plus / totalCount Else percentageLate = 0
            wsOutput.Cells(row_num, 1).Value = nameKey
            wsOutput.Cells(row_num, 2).Value = monthKey
            wsOutput.Cells(row_num, 3).Value = count7
            wsOutput.Cells(row_num, 4).Value = count14
            wsOutput.Cells(row_num, 5).Value = count14Plus
            wsOutput.Cells(row_num, 6).Value = percentageLate
            wsOutput.Cells(row_num, 6).NumberFormat = "0.00%"
            row_num = row_num + 1
        Next monthKey
    Next nameKey
    
    Dim chartDataStartRow As Long
    chartDataStartRow = 2
    Dim chartDataEndRow As Long
    chartDataEndRow = row_num - 1
    
    row_num = row_num + 2 ' Add a row for separation
    
    wsOutput.Cells(row_num, 1).Value = "Name"
    wsOutput.Cells(row_num, 2).Value = "Period"
    wsOutput.Cells(row_num, 3).Value = "<= 7 Days"
    wsOutput.Cells(row_num, 4).Value = "<= 14 Days"
    wsOutput.Cells(row_num, 5).Value = "> 14 Days"
    wsOutput.Cells(row_num, 6).Value = "Percentage Late"
    row_num = row_num + 1

    For Each nameKey In allNames.Keys
        Set allWeeks = CreateObject("Scripting.Dictionary")
        If dictWeekly7Days.Exists(nameKey) Then For Each subKey In dictWeekly7Days(nameKey).Keys: allWeeks(subKey) = True: Next
        If dictWeekly14Days.Exists(nameKey) Then For Each subKey In dictWeekly14Days(nameKey).Keys: allWeeks(subKey) = True: Next
        If dictWeekly14PlusDays.Exists(nameKey) Then For Each subKey In dictWeekly14PlusDays(nameKey).Keys: allWeeks(subKey) = True: Next
        For Each weekKey In allWeeks.Keys
            Dim weekCount7 As Long, weekCount14 As Long, weekCount14Plus As Long
            weekCount7 = IIf(dictWeekly7Days.Exists(nameKey) And dictWeekly7Days(nameKey).Exists(weekKey), dictWeekly7Days(nameKey)(weekKey), 0)
            weekCount14 = IIf(dictWeekly14Days.Exists(nameKey) And dictWeekly14Days(nameKey).Exists(weekKey), dictWeekly14Days(nameKey)(weekKey), 0)
            weekCount14Plus = IIf(dictWeekly14PlusDays.Exists(nameKey) And dictWeekly14PlusDays(nameKey).Exists(weekKey), dictWeekly14PlusDays(nameKey)(weekKey), 0)
            totalCount = weekCount7 + weekCount14 + weekCount14Plus
            If totalCount > 0 Then percentageLate = weekCount14Plus / totalCount Else percentageLate = 0
            wsOutput.Cells(row_num, 1).Value = nameKey
            wsOutput.Cells(row_num, 2).Value = Replace(weekKey, "_", " ")
            wsOutput.Cells(row_num, 3).Value = weekCount7
            wsOutput.Cells(row_num, 4).Value = weekCount14
            wsOutput.Cells(row_num, 5).Value = weekCount14Plus
            wsOutput.Cells(row_num, 6).Value = percentageLate
            wsOutput.Cells(row_num, 6).NumberFormat = "0.00%"
            row_num = row_num + 1
        Next weekKey
    Next nameKey

    row_num = row_num + 2 ' Add a row for separation

    wsOutput.Cells(row_num, 1).Value = "Name"
    wsOutput.Cells(row_num, 2).Value = "Prob. <= 7 Days"
    wsOutput.Cells(row_num, 3).Value = "Prob. <= 14 Days"
    wsOutput.Cells(row_num, 4).Value = "Prob. > 14 Days"
    row_num = row_num + 1
    
    For Each nameKey In dictTotalCounts.Keys
        Dim totalCountAll As Long
        totalCountAll = dictTotalCounts(nameKey)("Less7") + dictTotalCounts(nameKey)("Less14") + dictTotalCounts(nameKey)("Greater14")
        
        If totalCountAll > 0 Then
            wsOutput.Cells(row_num, 1).Value = nameKey
            wsOutput.Cells(row_num, 2).Value = dictTotalCounts(nameKey)("Less7") / totalCountAll
            wsOutput.Cells(row_num, 3).Value = dictTotalCounts(nameKey)("Less14") / totalCountAll
            wsOutput.Cells(row_num, 4).Value = dictTotalCounts(nameKey)("Greater14") / totalCountAll
            wsOutput.Cells(row_num, 2).NumberFormat = "0.00%"
            wsOutput.Cells(row_num, 3).NumberFormat = "0.00%"
            wsOutput.Cells(row_num, 4).NumberFormat = "0.00%"
        End If
        row_num = row_num + 1
    Next nameKey

    ' --- Create the Line Chart ---
    Dim chartRange As Range
    Dim ch As ChartObject
    Dim seriesCount As Long
    Dim chartData As Range
    
    ' Set the range for the chart data
    Set chartData = wsOutput.Range(wsOutput.Cells(chartDataStartRow, 1), wsOutput.Cells(chartDataEndRow, 6))
    
    ' Add a chart to the worksheet
    Set ch = wsOutput.ChartObjects.Add(Left:=50, Top:=wsOutput.Cells(row_num + 2, 1).Top, Width:=600, Height:=300)
    
    ' Set the chart properties
    With ch.Chart
        .ChartType = xlLine
        .SetSourceData Source:=chartData
        .HasTitle = True
        .ChartTitle.Text = "Percentage Late by Month"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Month"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Percentage Late"
        .Axes(xlValue).MaximumScale = 1 ' Set Y-axis to 0-100%
        .Axes(xlValue).MinimumScale = 0
    End With

    ' Customize each series
    For seriesCount = ch.Chart.SeriesCollection.Count To 1 Step -1
        ch.Chart.SeriesCollection(seriesCount).Delete
    Next seriesCount
    
    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")
    
    ' Group rows by name for charting
    Dim row_idx As Long
    For row_idx = chartDataStartRow To chartDataEndRow
        nameKey = wsOutput.Cells(row_idx, 1).Value
        If Not nameDict.Exists(nameKey) Then
            Set nameDict(nameKey) = CreateObject("Scripting.Dictionary")
        End If
        nameDict(nameKey)(wsOutput.Cells(row_idx, 2).Value) = wsOutput.Cells(row_idx, 6).Value
    Next row_idx
    
    ' Create a series for each name
    Dim seriesName As Variant
    For Each seriesName In nameDict.Keys
        Dim datesList As String, percentagesList As String
        Dim monthValues As Object
        Set monthValues = nameDict(seriesName)
        
        Dim monthKeyIter As Variant
        For Each monthKeyIter In monthValues.Keys
            If datesList = "" Then
                datesList = "'" & monthKeyIter & "'"
                percentagesList = CStr(monthValues(monthKeyIter))
            Else
                datesList = datesList & ",'" & monthKeyIter & "'"
                percentagesList = percentagesList & "," & CStr(monthValues(monthKeyIter))
            End If
        Next monthKeyIter
        
        With ch.Chart.SeriesCollection.NewSeries
            .Name = seriesName
            .XValues = "={" & datesList & "}"
            .Values = "={" & percentagesList & "}"
        End With
    Next seriesName
    
    wsOutput.Columns("A:H").AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Analysis complete and chart created. Results are in the 'DateRangeSummary' sheet.", vbInformation
    Exit Sub

Canceled:
    MsgBox "Operation canceled by user.", vbInformation
    Exit Sub
    
SheetError:
    MsgBox "Error: The sheet named 'Data' was not found. Please check your sheet name and try again.", vbCritical
End Sub
