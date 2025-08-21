Option Explicit

' Constants for thresholds and chart dimensions
Private Const DAYS_THRESHOLD_1 As Long = 7
Private Const DAYS_THRESHOLD_2 As Long = 14
Private Const CHART_WIDTH As Long = 600
Private Const CHART_HEIGHT As Long = 300

Function WeekOfMonth(ByVal d As Date) As Long
    ' Calculates the week number within a given month.
    ' Returns 0 if invalid date
    If d = 0 Then
        WeekOfMonth = 0
        Exit Function
    End If

    Dim firstDay As Date
    firstDay = DateSerial(Year(d), Month(d), 1)

    ' Calculate the week number based on the first day of the month
    WeekOfMonth = Int((Day(d) - 1 + Weekday(firstDay, vbMonday) - 1) / 7) + 1
End Function

Private Function IsValidDateRange(startDate As Date, endDate As Date) As Boolean
    ' Validates that dates are valid and in correct order
    IsValidDateRange = (startDate > 0 And endDate > 0 And startDate <= endDate)
End Function

Private Function GetOrCreateDictionary(dict As Object, key As String) As Object
    ' Helper function to get or create nested dictionary
    If Not dict.Exists(key) Then
        Set dict(key) = CreateObject("Scripting.Dictionary")
    End If
    Set GetOrCreateDictionary = dict(key)
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

    ' Variables for improved processing
    Dim dateInput As String
    Dim invalidRows As Long

    ' Store original application settings
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean

    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents

    ' Optimize Excel settings for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Check if Data sheet exists
    On Error GoTo SheetError
    Set wsData = ThisWorkbook.Sheets("Data")
    On Error GoTo 0

    ' Get date range from user with improved error handling
    On Error GoTo InvalidDateInput
    dateInput = Application.InputBox("Enter the start date (MM/DD/YYYY):", "Start Date", Type:=2)
    If dateInput = "False" Or dateInput = "" Then GoTo Canceled
    startDate = CDate(dateInput)

    dateInput = Application.InputBox("Enter the end date (MM/DD/YYYY):", "End Date", Type:=2)
    If dateInput = "False" Or dateInput = "" Then GoTo Canceled
    endDate = CDate(dateInput)
    On Error GoTo 0

    ' Validate date range
    If Not IsValidDateRange(startDate, endDate) Then
        MsgBox "Invalid date range. Start date must be before or equal to end date.", vbExclamation
        GoTo CleanUp
    End If

    ' Initialize dictionaries
    Set dictMonthly7Days = CreateObject("Scripting.Dictionary")
    Set dictMonthly14Days = CreateObject("Scripting.Dictionary")
    Set dictMonthly14PlusDays = CreateObject("Scripting.Dictionary")
    Set dictWeekly7Days = CreateObject("Scripting.Dictionary")
    Set dictWeekly14Days = CreateObject("Scripting.Dictionary")
    Set dictWeekly14PlusDays = CreateObject("Scripting.Dictionary")
    Set dictTotalCounts = CreateObject("Scripting.Dictionary")

    ' Find last row with data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    invalidRows = 0

    ' Process data with validation and progress indicator
    For currentRow = 2 To lastRow
        ' Update status bar for large datasets
        If currentRow Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & currentRow & " of " & lastRow & "..."
        End If

        ' Validate row data
        On Error Resume Next
        Dim cellA As Variant, cellB As Variant, cellC As Variant, cellD As Variant
        cellA = wsData.Cells(currentRow, "A").Value
        cellB = wsData.Cells(currentRow, "B").Value
        cellC = wsData.Cells(currentRow, "C").Value
        cellD = wsData.Cells(currentRow, "D").Value
        On Error GoTo 0

        ' Skip if invalid data or does not contain "pizza" in column D
        If cellA = "" Or Not IsDate(cellB) Or Not IsDate(cellC) Or Not IsEmpty(cellD) And InStr(1, CStr(cellD), "pizza", vbTextCompare) = 0 Then
            invalidRows = invalidRows + 1
            GoTo NextRow
        End If

        If CDate(cellB) >= startDate And CDate(cellB) <= endDate Then
            dateDiff = CDate(cellC) - CDate(cellB)
            nameKey = CStr(cellA)
            monthKey = Format(CDate(cellB), "YYYY-MM")
            weekNumber = WeekOfMonth(CDate(cellB))

            ' Initialize dictionary with default structure if it doesn't exist
            If Not dictTotalCounts.Exists(nameKey) Then
                Set dictTotalCounts(nameKey) = CreateObject("Scripting.Dictionary")
                dictTotalCounts(nameKey).Add "Less7", 0
                dictTotalCounts(nameKey).Add "Less14", 0
                dictTotalCounts(nameKey).Add "Greater14", 0
            End If

            If dateDiff <= DAYS_THRESHOLD_1 Then
                ' Process <= 7 days
                If Not dictMonthly7Days.Exists(nameKey) Then Set dictMonthly7Days(nameKey) = CreateObject("Scripting.Dictionary")
                If Not dictMonthly7Days(nameKey).Exists(monthKey) Then dictMonthly7Days(nameKey).Add monthKey, 0
                dictMonthly7Days(nameKey)(monthKey) = dictMonthly7Days(nameKey)(monthKey) + 1

                If Not dictWeekly7Days.Exists(nameKey) Then Set dictWeekly7Days(nameKey) = CreateObject("Scripting.Dictionary")
                Dim weekKey7 As String
                weekKey7 = monthKey & "_Week_" & weekNumber
                If Not dictWeekly7Days(nameKey).Exists(weekKey7) Then dictWeekly7Days(nameKey).Add weekKey7, 0
                dictWeekly7Days(nameKey)(weekKey7) = dictWeekly7Days(nameKey)(weekKey7) + 1

                dictTotalCounts(nameKey)("Less7") = dictTotalCounts(nameKey)("Less7") + 1

            ElseIf dateDiff <= DAYS_THRESHOLD_2 Then
                ' Process <= 14 days
                If Not dictMonthly14Days.Exists(nameKey) Then Set dictMonthly14Days(nameKey) = CreateObject("Scripting.Dictionary")
                If Not dictMonthly14Days(nameKey).Exists(monthKey) Then dictMonthly14Days(nameKey).Add monthKey, 0
                dictMonthly14Days(nameKey)(monthKey) = dictMonthly14Days(nameKey)(monthKey) + 1

                If Not dictWeekly14Days.Exists(nameKey) Then Set dictWeekly14Days(nameKey) = CreateObject("Scripting.Dictionary")
                Dim weekKey14 As String
                weekKey14 = monthKey & "_Week_" & weekNumber
                If Not dictWeekly14Days(nameKey).Exists(weekKey14) Then dictWeekly14Days(nameKey).Add weekKey14, 0
                dictWeekly14Days(nameKey)(weekKey14) = dictWeekly14Days(nameKey)(weekKey14) + 1

                dictTotalCounts(nameKey)("Less14") = dictTotalCounts(nameKey)("Less14") + 1

            Else
                ' Process > 14 days
                If Not dictMonthly14PlusDays.Exists(nameKey) Then Set dictMonthly14PlusDays(nameKey) = CreateObject("Scripting.Dictionary")
                If Not dictMonthly14PlusDays(nameKey).Exists(monthKey) Then dictMonthly14PlusDays(nameKey).Add monthKey, 0
                dictMonthly14PlusDays(nameKey)(monthKey) = dictMonthly14PlusDays(nameKey)(monthKey) + 1

                If Not dictWeekly14PlusDays.Exists(nameKey) Then Set dictWeekly14PlusDays(nameKey) = CreateObject("Scripting.Dictionary")
                Dim weekKeyPlus As String
                weekKeyPlus = monthKey & "_Week_" & weekNumber
                If Not dictWeekly14PlusDays(nameKey).Exists(weekKeyPlus) Then dictWeekly14PlusDays(nameKey).Add weekKeyPlus, 0
                dictWeekly14PlusDays(nameKey)(weekKeyPlus) = dictWeekly14PlusDays(nameKey)(weekKeyPlus) + 1

                dictTotalCounts(nameKey)("Greater14") = dictTotalCounts(nameKey)("Greater14") + 1
            End If
        End If
NextRow:
    Next currentRow

    ' Clear status bar
    Application.StatusBar = False

    ' Create new output sheet or clear existing one
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("DateRangeSummary")
    If Not wsOutput Is Nothing Then
        Application.DisplayAlerts = False
        wsOutput.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsData)
    wsOutput.Name = "DateRangeSummary"

    ' Delete any existing charts in the new sheet
    Dim co As ChartObject
    For Each co In wsOutput.ChartObjects
        co.Delete
    Next co

    ' --- Output Monthly/Weekly Data ---
    Dim allNames As Object, allMonths As Object, allWeeks As Object
    Set allNames = CreateObject("Scripting.Dictionary")

    ' Declare variables properly
    Dim nameVar As Variant, subKeyVar As Variant, monthVar As Variant, weekVar As Variant

    ' Collect all unique names
    For Each nameVar In dictMonthly7Days.Keys: allNames(nameVar) = True: Next
    For Each nameVar In dictMonthly14Days.Keys: allNames(nameVar) = True: Next
    For Each nameVar In dictMonthly14PlusDays.Keys: allNames(nameVar) = True: Next
    For Each nameVar In dictWeekly7Days.Keys: allNames(nameVar) = True: Next
    For Each nameVar In dictWeekly14Days.Keys: allNames(nameVar) = True: Next
    For Each nameVar In dictWeekly14PlusDays.Keys: allNames(nameVar) = True: Next

    Dim outputRow As Long
    outputRow = 1

    ' Monthly data headers
    wsOutput.Cells(outputRow, 1).Value = "Name"
    wsOutput.Cells(outputRow, 2).Value = "Period"
    wsOutput.Cells(outputRow, 3).Value = "<= 7 Days"
    wsOutput.Cells(outputRow, 4).Value = "<= 14 Days"
    wsOutput.Cells(outputRow, 5).Value = "> 14 Days"
    wsOutput.Cells(outputRow, 6).Value = "Percentage Late"

    ' Format headers
    With wsOutput.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    outputRow = outputRow + 1

    ' Process monthly data - collect all data first
    Dim monthlyData As Object
    Set monthlyData = CreateObject("Scripting.Dictionary")

    For Each nameVar In allNames.Keys
        nameKey = CStr(nameVar)
        Set allMonths = CreateObject("Scripting.Dictionary")

        If dictMonthly7Days.Exists(nameKey) Then
            For Each subKeyVar In dictMonthly7Days(nameKey).Keys: allMonths(subKeyVar) = True: Next
        End If
        If dictMonthly14Days.Exists(nameKey) Then
            For Each subKeyVar In dictMonthly14Days(nameKey).Keys: allMonths(subKeyVar) = True: Next
        End If
        If dictMonthly14PlusDays.Exists(nameKey) Then
            For Each subKeyVar In dictMonthly14PlusDays(nameKey).Keys: allMonths(subKeyVar) = True: Next
        End If

        For Each monthVar In allMonths.Keys
            monthKey = CStr(monthVar)
            Dim count7 As Long, count14 As Long, count14Plus As Long, totalCount As Long, percentageLate As Double
            count7 = 0: count14 = 0: count14Plus = 0

            If dictMonthly7Days.Exists(nameKey) Then
                If dictMonthly7Days(nameKey).Exists(monthKey) Then count7 = dictMonthly7Days(nameKey)(monthKey)
            End If
            If dictMonthly14Days.Exists(nameKey) Then
                If dictMonthly14Days(nameKey).Exists(monthKey) Then count14 = dictMonthly14Days(nameKey)(monthKey)
            End If
            If dictMonthly14PlusDays.Exists(nameKey) Then
                If dictMonthly14PlusDays(nameKey).Exists(monthKey) Then count14Plus = dictMonthly14PlusDays(nameKey)(monthKey)
            End If

            totalCount = count7 + count14 + count14Plus
            If totalCount > 0 Then percentageLate = count14Plus / totalCount Else percentageLate = 0

            ' Store data in dictionary organized by month
            If Not monthlyData.Exists(monthKey) Then
                Set monthlyData(monthKey) = CreateObject("Scripting.Dictionary")
            End If

            ' Create unique key with percentage for sorting
            Dim dataKey As String
            dataKey = Format(1 - percentageLate, "0.00000") & "_" & nameKey ' Inverse for descending sort

            Set monthlyData(monthKey)(dataKey) = CreateObject("Scripting.Dictionary")
            monthlyData(monthKey)(dataKey)("Name") = nameKey
            monthlyData(monthKey)(dataKey)("Count7") = count7
            monthlyData(monthKey)(dataKey)("Count14") = count14
            monthlyData(monthKey)(dataKey)("Count14Plus") = count14Plus
            monthlyData(monthKey)(dataKey)("PercentageLate") = percentageLate
        Next monthVar
    Next nameVar

    ' Output monthly data grouped by month and sorted by percentage late
    Dim monthKeys() As String
    Dim monthCount As Long
    monthCount = monthlyData.Count

    If monthCount > 0 Then
        ReDim monthKeys(0 To monthCount - 1)
        Dim idx As Long
        idx = 0
        For Each monthVar In monthlyData.Keys
            monthKeys(idx) = CStr(monthVar)
            idx = idx + 1
        Next monthVar

        ' Sort months
        Dim tempMonth As String
        Dim j As Long
        For idx = 0 To monthCount - 2
            For j = idx + 1 To monthCount - 1
                If monthKeys(idx) > monthKeys(j) Then
                    tempMonth = monthKeys(idx)
                    monthKeys(idx) = monthKeys(j)
                    monthKeys(j) = tempMonth
                End If
            Next j
        Next idx

        ' Output each month's data
        For idx = 0 To monthCount - 1
            monthKey = monthKeys(idx)

            ' Add month header
            wsOutput.Cells(outputRow, 1).Value = "Month: " & monthKey
            With wsOutput.Range(wsOutput.Cells(outputRow, 1), wsOutput.Cells(outputRow, 6))
                .Merge
                .Font.Bold = True
                .Font.Size = 12
                .Interior.Color = RGB(180, 180, 180)
                .HorizontalAlignment = xlCenter
            End With
            outputRow = outputRow + 1

            ' Add column headers for this month
            wsOutput.Cells(outputRow, 1).Value = "Name"
            wsOutput.Cells(outputRow, 2).Value = "Period"
            wsOutput.Cells(outputRow, 3).Value = "<= 7 Days"
            wsOutput.Cells(outputRow, 4).Value = "<= 14 Days"
            wsOutput.Cells(outputRow, 5).Value = "> 14 Days"
            wsOutput.Cells(outputRow, 6).Value = "Percentage Late"
            With wsOutput.Range("A" & outputRow & ":F" & outputRow)
                .Font.Bold = True
                .Interior.Color = RGB(220, 220, 220)
            End With
            outputRow = outputRow + 1

            ' Variables for month totals
            Dim monthTotal7 As Long, monthTotal14 As Long, monthTotal14Plus As Long
            monthTotal7 = 0: monthTotal14 = 0: monthTotal14Plus = 0

            ' Get sorted keys for this month
            Dim sortedKeys() As String
            Dim keyCount As Long
            keyCount = monthlyData(monthKey).Count

            If keyCount > 0 Then
                ReDim sortedKeys(0 To keyCount - 1)
                Dim keyIdx As Long
                keyIdx = 0
                Dim dataKeyVar As Variant
                For Each dataKeyVar In monthlyData(monthKey).Keys
                    sortedKeys(keyIdx) = CStr(dataKeyVar)
                    keyIdx = keyIdx + 1
                Next dataKeyVar

                ' Sort keys (already formatted for proper sorting)
                Dim tempKey As String
                Dim k As Long
                For keyIdx = 0 To keyCount - 2
                    For k = keyIdx + 1 To keyCount - 1
                        If sortedKeys(keyIdx) > sortedKeys(k) Then
                            tempKey = sortedKeys(keyIdx)
                            sortedKeys(keyIdx) = sortedKeys(k)
                            sortedKeys(k) = tempKey
                        End If
                    Next k
                Next keyIdx

                ' Output sorted data for this month
                For keyIdx = 0 To keyCount - 1
                    Dim dataDict As Object
                    Set dataDict = monthlyData(monthKey)(sortedKeys(keyIdx))

                    wsOutput.Cells(outputRow, 1).Value = dataDict("Name")
                    wsOutput.Cells(outputRow, 2).Value = monthKey
                    wsOutput.Cells(outputRow, 3).Value = dataDict("Count7")
                    wsOutput.Cells(outputRow, 4).Value = dataDict("Count14")
                    wsOutput.Cells(outputRow, 5).Value = dataDict("Count14Plus")
                    wsOutput.Cells(outputRow, 6).Value = dataDict("PercentageLate")
                    wsOutput.Cells(outputRow, 6).NumberFormat = "0.00%"

                    ' Add to totals
                    monthTotal7 = monthTotal7 + dataDict("Count7")
                    monthTotal14 = monthTotal14 + dataDict("Count14")
                    monthTotal14Plus = monthTotal14Plus + dataDict("Count14Plus")

                    ' Highlight high late percentages
                    If dataDict("PercentageLate") > 0.5 Then
                        wsOutput.Cells(outputRow, 6).Font.Color = RGB(255, 0, 0)
                    ElseIf dataDict("PercentageLate") > 0.25 Then
                        wsOutput.Cells(outputRow, 6).Font.Color = RGB(255, 128, 0)
                    End If

                    outputRow = outputRow + 1
                Next keyIdx

                ' Add totals row
                wsOutput.Cells(outputRow, 1).Value = "TOTAL"
                wsOutput.Cells(outputRow, 2).Value = monthKey
                wsOutput.Cells(outputRow, 3).Value = monthTotal7
                wsOutput.Cells(outputRow, 4).Value = monthTotal14
                wsOutput.Cells(outputRow, 5).Value = monthTotal14Plus
                Dim monthGrandTotal As Long
                monthGrandTotal = monthTotal7 + monthTotal14 + monthTotal14Plus
                If monthGrandTotal > 0 Then
                    wsOutput.Cells(outputRow, 6).Value = monthTotal14Plus / monthGrandTotal
                Else
                    wsOutput.Cells(outputRow, 6).Value = 0
                End If
                wsOutput.Cells(outputRow, 6).NumberFormat = "0.00%"

                ' Format totals row
                With wsOutput.Range("A" & outputRow & ":F" & outputRow)
                    .Font.Bold = True
                    .Interior.Color = RGB(240, 240, 240)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With

                outputRow = outputRow + 1
            End If

            ' Add empty row after each month group
            outputRow = outputRow + 1
        Next idx
    End If

    ' Clean up
    Set monthlyData = Nothing

    Dim chartDataStartRow As Long
    chartDataStartRow = 2
    Dim chartDataEndRow As Long
    chartDataEndRow = outputRow - 1

    outputRow = outputRow + 2 ' Add separation

    ' Weekly data headers
    wsOutput.Cells(outputRow, 1).Value = "Name"
    wsOutput.Cells(outputRow, 2).Value = "Period"
    wsOutput.Cells(outputRow, 3).Value = "<= 7 Days"
    wsOutput.Cells(outputRow, 4).Value = "<= 14 Days"
    wsOutput.Cells(outputRow, 5).Value = "> 14 Days"
    wsOutput.Cells(outputRow, 6).Value = "Percentage Late"

    ' Format headers
    With wsOutput.Range("A" & outputRow & ":F" & outputRow)
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    outputRow = outputRow + 1

    ' Process weekly data - collect and organize by month/week
    Dim weeklyData As Object
    Set weeklyData = CreateObject("Scripting.Dictionary")

    For Each nameVar In allNames.Keys
        nameKey = CStr(nameVar)
        Set allWeeks = CreateObject("Scripting.Dictionary")

        If dictWeekly7Days.Exists(nameKey) Then
            For Each subKeyVar In dictWeekly7Days(nameKey).Keys: allWeeks(subKeyVar) = True: Next
        End If
        If dictWeekly14Days.Exists(nameKey) Then
            For Each subKeyVar In dictWeekly14Days(nameKey).Keys: allWeeks(subKeyVar) = True: Next
        End If
        If dictWeekly14PlusDays.Exists(nameKey) Then
            For Each subKeyVar In dictWeekly14PlusDays(nameKey).Keys: allWeeks(subKeyVar) = True: Next
        End If

        For Each weekVar In allWeeks.Keys
            Dim weekKey As String
            weekKey = CStr(weekVar)

            ' Extract month from week key (format: YYYY-MM_Week_N)
            Dim weekMonth As String
            weekMonth = Left(weekKey, 7) ' Get YYYY-MM part

            Dim weekCount7 As Long, weekCount14 As Long, weekCount14Plus As Long
            weekCount7 = 0: weekCount14 = 0: weekCount14Plus = 0

            If dictWeekly7Days.Exists(nameKey) Then
                If dictWeekly7Days(nameKey).Exists(weekKey) Then weekCount7 = dictWeekly7Days(nameKey)(weekKey)
            End If
            If dictWeekly14Days.Exists(nameKey) Then
                If dictWeekly14Days(nameKey).Exists(weekKey) Then weekCount14 = dictWeekly14Days(nameKey)(weekKey)
            End If
            If dictWeekly14PlusDays.Exists(nameKey) Then
                If dictWeekly14PlusDays(nameKey).Exists(weekKey) Then weekCount14Plus = dictWeekly14PlusDays(nameKey)(weekKey)
            End If

            totalCount = weekCount7 + weekCount14 + weekCount14Plus
            If totalCount > 0 Then percentageLate = weekCount14Plus / totalCount Else percentageLate = 0

            ' Store data organized by month
            If Not weeklyData.Exists(weekMonth) Then
                Set weeklyData(weekMonth) = CreateObject("Scripting.Dictionary")
            End If

            ' Create unique key with percentage for sorting
            Dim weekDataKey As String
            weekDataKey = Format(1 - percentageLate, "0.00000") & "_" & weekKey & "_" & nameKey

            Set weeklyData(weekMonth)(weekDataKey) = CreateObject("Scripting.Dictionary")
            weeklyData(weekMonth)(weekDataKey)("Name") = nameKey
            weeklyData(weekMonth)(weekDataKey)("Week") = Replace(weekKey, "_", " ")
            weeklyData(weekMonth)(weekDataKey)("Count7") = weekCount7
            weeklyData(weekMonth)(weekDataKey)("Count14") = weekCount14
            weeklyData(weekMonth)(weekDataKey)("Count14Plus") = weekCount14Plus
            weeklyData(weekMonth)(weekDataKey)("PercentageLate") = percentageLate
        Next weekVar
    Next nameVar

    ' Output weekly data grouped by month and sorted by percentage late
    Dim weekMonthKeys() As String
    Dim weekMonthCount As Long
    weekMonthCount = weeklyData.Count

    If weekMonthCount > 0 Then
        ReDim weekMonthKeys(0 To weekMonthCount - 1)
        idx = 0
        For Each monthVar In weeklyData.Keys
            weekMonthKeys(idx) = CStr(monthVar)
            idx = idx + 1
        Next monthVar

        ' Sort months
        For idx = 0 To weekMonthCount - 2
            For j = idx + 1 To weekMonthCount - 1
                If weekMonthKeys(idx) > weekMonthKeys(j) Then
                    tempMonth = weekMonthKeys(idx)
                    weekMonthKeys(idx) = weekMonthKeys(j)
                    weekMonthKeys(j) = tempMonth
                End If
            Next j
        Next idx

        ' Output each month's weekly data
        For idx = 0 To weekMonthCount - 1
            Dim weekMonthKey As String
            weekMonthKey = weekMonthKeys(idx)

            ' Add month header for weekly data
            wsOutput.Cells(outputRow, 1).Value = "Month: " & weekMonthKey & " (Weekly Breakdown)"
            With wsOutput.Range(wsOutput.Cells(outputRow, 1), wsOutput.Cells(outputRow, 6))
                .Merge
                .Font.Bold = True
                .Font.Size = 12
                .Interior.Color = RGB(180, 180, 180)
                .HorizontalAlignment = xlCenter
            End With
            outputRow = outputRow + 1

            ' Add column headers for this month's weekly data
            wsOutput.Cells(outputRow, 1).Value = "Name"
            wsOutput.Cells(outputRow, 2).Value = "Period"
            wsOutput.Cells(outputRow, 3).Value = "<= 7 Days"
            wsOutput.Cells(outputRow, 4).Value = "<= 14 Days"
            wsOutput.Cells(outputRow, 5).Value = "> 14 Days"
            wsOutput.Cells(outputRow, 6).Value = "Percentage Late"
            With wsOutput.Range("A" & outputRow & ":F" & outputRow)
                .Font.Bold = True
                .Interior.Color = RGB(220, 220, 220)
            End With
            outputRow = outputRow + 1

            ' Variables for weekly totals
            Dim weekTotal7 As Long, weekTotal14 As Long, weekTotal14Plus As Long
            weekTotal7 = 0: weekTotal14 = 0: weekTotal14Plus = 0

            ' Get sorted keys for this month's weeks
            Dim weekSortedKeys() As String
            Dim weekKeyCount As Long
            weekKeyCount = weeklyData(weekMonthKey).Count

            If weekKeyCount > 0 Then
                ReDim weekSortedKeys(0 To weekKeyCount - 1)
                Dim weekKeyIdx As Long
                weekKeyIdx = 0
                For Each dataKeyVar In weeklyData(weekMonthKey).Keys
                    weekSortedKeys(weekKeyIdx) = CStr(dataKeyVar)
                    weekKeyIdx = weekKeyIdx + 1
                Next dataKeyVar

                ' Sort keys
                For weekKeyIdx = 0 To weekKeyCount - 2
                    For k = weekKeyIdx + 1 To weekKeyCount - 1
                        If weekSortedKeys(weekKeyIdx) > weekSortedKeys(k) Then
                            tempKey = weekSortedKeys(weekKeyIdx)
                            weekSortedKeys(weekKeyIdx) = weekSortedKeys(k)
                            weekSortedKeys(k) = tempKey
                        End If
                    Next k
                Next weekKeyIdx

                ' Output sorted weekly data for this month
                For weekKeyIdx = 0 To weekKeyCount - 1
                    Dim weekDataDict As Object
                    Set weekDataDict = weeklyData(weekMonthKey)(weekSortedKeys(weekKeyIdx))

                    wsOutput.Cells(outputRow, 1).Value = weekDataDict("Name")
                    wsOutput.Cells(outputRow, 2).Value = weekDataDict("Week")
                    wsOutput.Cells(outputRow, 3).Value = weekDataDict("Count7")
                    wsOutput.Cells(outputRow, 4).Value = weekDataDict("Count14")
                    wsOutput.Cells(outputRow, 5).Value = weekDataDict("Count14Plus")
                    wsOutput.Cells(outputRow, 6).Value = weekDataDict("PercentageLate")
                    wsOutput.Cells(outputRow, 6).NumberFormat = "0.00%"

                    ' Add to totals
                    weekTotal7 = weekTotal7 + weekDataDict("Count7")
                    weekTotal14 = weekTotal14 + weekDataDict("Count14")
                    weekTotal14Plus = weekTotal14Plus + weekDataDict("Count14Plus")

                    ' Highlight high late percentages
                    If weekDataDict("PercentageLate") > 0.5 Then
                        wsOutput.Cells(outputRow, 6).Font.Color = RGB(255, 0, 0)
                    ElseIf weekDataDict("PercentageLate") > 0.25 Then
                        wsOutput.Cells(outputRow, 6).Font.Color = RGB(255, 128, 0)
                    End If

                    outputRow = outputRow + 1
                Next weekKeyIdx

                ' Add totals row
                wsOutput.Cells(outputRow, 1).Value = "TOTAL"
                wsOutput.Cells(outputRow, 2).Value = weekMonthKey
                wsOutput.Cells(outputRow, 3).Value = weekTotal7
                wsOutput.Cells(outputRow, 4).Value = weekTotal14
                wsOutput.Cells(outputRow, 5).Value = weekTotal14Plus
                Dim weekGrandTotal As Long
                weekGrandTotal = weekTotal7 + weekTotal14 + weekTotal14Plus
                If weekGrandTotal > 0 Then
                    wsOutput.Cells(outputRow, 6).Value = weekTotal14Plus / weekGrandTotal
                Else
                    wsOutput.Cells(outputRow, 6).Value = 0
                End If
                wsOutput.Cells(outputRow, 6).NumberFormat = "0.00%"

                ' Format totals row
                With wsOutput.Range("A" & outputRow & ":F" & outputRow)
                    .Font.Bold = True
                    .Interior.Color = RGB(240, 240, 240)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With

                outputRow = outputRow + 1
            End If

            ' Add empty row after each month group
            outputRow = outputRow + 1
        Next idx
    End If

    ' Clean up
    Set weeklyData = Nothing

    outputRow = outputRow + 2 ' Add separation

    ' Probability summary headers
    wsOutput.Cells(outputRow, 1).Value = "Name"
    wsOutput.Cells(outputRow, 2).Value = "Prob. <= 7 Days"
    wsOutput.Cells(outputRow, 3).Value = "Prob. <= 14 Days"
    wsOutput.Cells(outputRow, 4).Value = "Prob. > 14 Days"

    ' Format headers
    With wsOutput.Range("A" & outputRow & ":D" & outputRow)
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    outputRow = outputRow + 1

    ' Process probability summary
    For Each nameVar In dictTotalCounts.Keys
        nameKey = CStr(nameVar)
        Dim totalCountAll As Long
        totalCountAll = dictTotalCounts(nameKey)("Less7") + dictTotalCounts(nameKey)("Less14") + dictTotalCounts(nameKey)("Greater14")

        If totalCountAll > 0 Then
            wsOutput.Cells(outputRow, 1).Value = nameKey
            wsOutput.Cells(outputRow, 2).Value = dictTotalCounts(nameKey)("Less7") / totalCountAll
            wsOutput.Cells(outputRow, 3).Value = dictTotalCounts(nameKey)("Less14") / totalCountAll
            wsOutput.Cells(outputRow, 4).Value = dictTotalCounts(nameKey)("Greater14") / totalCountAll
            wsOutput.Cells(outputRow, 2).NumberFormat = "0.00%"
            wsOutput.Cells(outputRow, 3).NumberFormat = "0.00%"
            wsOutput.Cells(outputRow, 4).NumberFormat = "0.00%"
        End If
        outputRow = outputRow + 1
    Next nameVar

    ' --- Create the Line Chart ---
    If chartDataEndRow >= chartDataStartRow Then
        On Error Resume Next
        Dim chartRange As Range
        Dim ch As ChartObject
        Dim seriesCount As Long
        Dim chartData As Range

        ' Add a chart to the worksheet
        Set ch = wsOutput.ChartObjects.Add(Left:=50, Top:=wsOutput.Cells(outputRow + 2, 1).Top, Width:=CHART_WIDTH, Height:=CHART_HEIGHT)

        If Not ch Is Nothing Then
            ' Set the chart properties
            With ch.Chart
                .ChartType = xlLine
                .HasTitle = True
                .ChartTitle.Text = "Percentage Late by Month"
                .Axes(xlCategory, xlPrimary).HasTitle = True
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Month"
                .Axes(xlValue, xlPrimary).HasTitle = True
                .Axes(xlValue, xlPrimary).AxisTitle.Text = "Percentage Late"
                .Axes(xlValue).MaximumScale = 1 ' Set Y-axis to 0-100%
                .Axes(xlValue).MinimumScale = 0
            End With

            ' Clear any default series
            Do While ch.Chart.SeriesCollection.Count > 0
                ch.Chart.SeriesCollection(1).Delete
            Loop

            ' Create a dictionary to collect data by name across all months
            Dim chartDataDict As Object
            Set chartDataDict = CreateObject("Scripting.Dictionary")

            ' Collect all unique names and months for the chart
            Dim uniqueNames As Object
            Set uniqueNames = CreateObject("Scripting.Dictionary")
            Dim uniqueMonths As Object
            Set uniqueMonths = CreateObject("Scripting.Dictionary")

            ' Read through the monthly data section to build chart data
            Dim chartRow As Long
            For chartRow = 2 To chartDataEndRow
                ' Skip header rows, total rows, and empty rows
                If wsOutput.Cells(chartRow, 1).Value <> "" And _
                   wsOutput.Cells(chartRow, 1).Value <> "Name" And _
                   wsOutput.Cells(chartRow, 1).Value <> "TOTAL" And _
                   InStr(wsOutput.Cells(chartRow, 1).Value, "Month:") = 0 Then

                    Dim chartName As String
                    Dim chartMonth As String
                    Dim chartPercentage As Double

                    chartName = wsOutput.Cells(chartRow, 1).Value
                    chartMonth = wsOutput.Cells(chartRow, 2).Value
                    chartPercentage = wsOutput.Cells(chartRow, 6).Value

                    ' Store unique names and months
                    uniqueNames(chartName) = True
                    uniqueMonths(chartMonth) = True

                    ' Store data in nested dictionary
                    If Not chartDataDict.Exists(chartName) Then
                        Set chartDataDict(chartName) = CreateObject("Scripting.Dictionary")
                    End If
                    chartDataDict(chartName)(chartMonth) = chartPercentage
                End If
            Next chartRow

            ' Sort months for x-axis
            Dim sortedMonths() As String
            Dim monthIdx As Long
            ReDim sortedMonths(0 To uniqueMonths.Count - 1)
            monthIdx = 0
            Dim monthItem As Variant
            For Each monthItem In uniqueMonths.Keys
                sortedMonths(monthIdx) = CStr(monthItem)
                monthIdx = monthIdx + 1
            Next monthItem

            ' Sort the months array
            Dim m As Long, n As Long
            For m = 0 To UBound(sortedMonths) - 1
                For n = m + 1 To UBound(sortedMonths)
                    If sortedMonths(m) > sortedMonths(n) Then
                        Dim tempSortMonth As String
                        tempSortMonth = sortedMonths(m)
                        sortedMonths(m) = sortedMonths(n)
                        sortedMonths(n) = tempSortMonth
                    End If
                Next n
            Next m

            ' Create a series for each unique name
            Dim nameItem As Variant
            For Each nameItem In uniqueNames.Keys
                Dim seriesName As String
                seriesName = CStr(nameItem)

                If chartDataDict.Exists(seriesName) Then
                    ' Build arrays for this series
                    Dim seriesValues() As Double
                    ReDim seriesValues(0 To UBound(sortedMonths))

                    Dim hasData As Boolean
                    hasData = False

                    For monthIdx = 0 To UBound(sortedMonths)
                        If chartDataDict(seriesName).Exists(sortedMonths(monthIdx)) Then
                            seriesValues(monthIdx) = chartDataDict(seriesName)(sortedMonths(monthIdx))
                            hasData = True
                        Else
                            ' No data for this month - use 0 or skip
                            seriesValues(monthIdx) = 0
                        End If
                    Next monthIdx

                    ' Only add series if it has data
                    If hasData Then
                        With ch.Chart.SeriesCollection.NewSeries
                            .Name = seriesName
                            .XValues = sortedMonths
                            .Values = seriesValues
                            .MarkerStyle = xlMarkerStyleCircle
                            .MarkerSize = 5
                        End With
                    End If
                End If
            Next nameItem

            ' Format the chart
            With ch.Chart
                .Legend.Position = xlLegendPositionRight
                .PlotArea.Border.LineStyle = xlContinuous
                .PlotArea.Border.Color = RGB(200, 200, 200)
            End With

            ' Clean up chart dictionary
            Set chartDataDict = Nothing
            Set uniqueNames = Nothing
            Set uniqueMonths = Nothing
        End If
        On Error GoTo 0
    End If

    ' Auto-fit columns
    wsOutput.Columns("A:H").AutoFit

    ' Display summary message
    Dim successMsg As String
    successMsg = "Analysis complete!" & vbCrLf & vbCrLf
    successMsg = successMsg & "Date Range: " & Format(startDate, "MM/DD/YYYY") & " to " & Format(endDate, "MM/DD/YYYY") & vbCrLf
    successMsg = successMsg & "Rows Processed: " & (lastRow - 1 - invalidRows) & vbCrLf
    If invalidRows > 0 Then
        successMsg = successMsg & "Invalid Rows Skipped: " & invalidRows & vbCrLf
    End If
    successMsg = successMsg & vbCrLf & "Results are in the 'DateRangeSummary' sheet."

    MsgBox successMsg, vbInformation, "Analysis Complete"

    GoTo CleanUp

' Error Handlers
InvalidDateInput:
    MsgBox "Invalid date format. Please enter dates in MM/DD/YYYY format.", vbExclamation, "Date Error"
    GoTo CleanUp

Canceled:
    MsgBox "Operation canceled by user.", vbInformation, "Canceled"
    GoTo CleanUp

SheetError:
    MsgBox "Error: The sheet named 'Data' was not found. Please ensure you have a sheet named 'Data' with the following columns:" & vbCrLf & _
           "Column A: Names" & vbCrLf & _
           "Column B: Start Dates" & vbCrLf & _
           "Column C: End Dates", vbCritical, "Sheet Not Found"
    GoTo CleanUp

CleanUp:
    ' Clean up dictionary objects
    Set dictMonthly7Days = Nothing
    Set dictMonthly14Days = Nothing
    Set dictMonthly14PlusDays = Nothing
    Set dictWeekly7Days = Nothing
    Set dictWeekly14Days = Nothing
    Set dictWeekly14PlusDays = Nothing
    Set dictTotalCounts = Nothing
    Set allNames = Nothing
    Set allMonths = Nothing
    Set allWeeks = Nothing

    ' Clear status bar
    Application.StatusBar = False

    ' Restore original application settings
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
End Sub

