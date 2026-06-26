Sub DataCleanup()
    Dim dataWs As Worksheet
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim tempWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    Dim cellVal As String
    Dim found As Boolean
    Dim s1Name As String
    Dim s2Name As String
    s1Name = "Anomaly Count Per Disposition"
    s2Name = "Defect Code Analysis"
    Set dataWs = ActiveSheet
    If dataWs.Name = s1Name Or dataWs.Name = s2Name Then
        For Each tempWs In ThisWorkbook.Worksheets
            If tempWs.Name <> s1Name And tempWs.Name <> s2Name Then
                Set dataWs = tempWs
                Exit For
            End If
        Next tempWs
    End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    lastRow = dataWs.Cells(dataWs.Rows.Count, "H").End(xlUp).Row
    For i = lastRow To 4 Step -1
        If SafeText(dataWs.Cells(i, "H")) <> "CLOSED" Then
            dataWs.Rows(i).Delete
        End If
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, "U").End(xlUp).Row
    For i = lastRow To 4 Step -1
        If Left(SafeText(dataWs.Cells(i, "U")), 1) <> "4" Then
            dataWs.Rows(i).Delete
        End If
    Next i
    dataWs.Columns(23).Insert Shift:=xlToRight
    dataWs.Cells(3, 23).Value = "Extracted"
    lastRow = dataWs.Cells(dataWs.Rows.Count, "V").End(xlUp).Row
    If lastRow >= 4 Then
        dataWs.Range(dataWs.Cells(4, 23), dataWs.Cells(lastRow, 23)).Formula = _
            "=LEFT(V4,FIND(""-"",V4&""-"")-1)"
    End If
    Application.Calculate
    lastRow = dataWs.Cells(dataWs.Rows.Count, 23).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(Left(SafeText(dataWs.Cells(i, 23)), 9)) = "NO DEFECT" Then
            dataWs.Rows(i).Delete
        End If
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, 24).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(SafeText(dataWs.Cells(i, 24))) = "PROMOTE TO HAR" Then
            dataWs.Rows(i).Delete
        End If
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, 24).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(SafeText(dataWs.Cells(i, 24))) = "NDF" Then
            dataWs.Rows(i).Delete
        End If
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    Dim suppliers() As String
    Dim supplierCount As Integer
    supplierCount = 0
    For i = 4 To lastRow
        cellVal = SafeText(dataWs.Cells(i, "A"))
        If cellVal <> "" Then
            found = False
            For j = 1 To supplierCount
                If suppliers(j) = cellVal Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                supplierCount = supplierCount + 1
                ReDim Preserve suppliers(1 To supplierCount)
                suppliers(supplierCount) = cellVal
            End If
        End If
    Next i
    If supplierCount = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "No supplier data found in column A after cleanup.", vbExclamation
        Exit Sub
    End If
    Application.DisplayAlerts = False
    For Each tempWs In ThisWorkbook.Worksheets
        If tempWs.Name = s1Name Then
            tempWs.Delete
            Exit For
        End If
    Next tempWs
    Application.DisplayAlerts = True
    Set ws1 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error Resume Next
    ws1.Name = s1Name
    If ws1.Name <> s1Name Then ws1.Name = Left(s1Name, 31)
    On Error GoTo 0
    With ws1.Cells(1, 1)
        .Value = "Supplier:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    For i = 1 To supplierCount
        ws1.Cells(i, 26).Value = suppliers(i)
    Next i
    ws1.Cells(1, 2).Value = suppliers(1)
    With ws1.Cells(1, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=$Z$1:$Z$" & supplierCount
        .InCellDropdown = True
    End With
    With ws1.Cells(1, 2)
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End With
    With ws1.Cells(3, 1)
        .Value = "Disposition Code"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    With ws1.Cells(3, 2)
        .Value = "Count"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws1.Columns("A:B").AutoFit
    ws1.Columns("Z").ColumnWidth = 0.1
    Dim btn1 As Button
    Set btn1 = ws1.Buttons.Add(ws1.Cells(1, 3).Left, ws1.Cells(1, 3).Top, 100, 22)
    btn1.Caption = "Refresh"
    btn1.OnAction = "RefreshSheet1"
    Dim d1Codes() As String
    Dim d1Counts() As Long
    Dim d1Count As Integer
    Dim selSup As String
    Dim dVal As String
    d1Count = 0
    selSup = suppliers(1)
    For i = 4 To lastRow
        If SafeText(dataWs.Cells(i, "A")) = selSup Then
            dVal = SafeText(dataWs.Cells(i, 24))
            If dVal <> "" Then
                found = False
                For j = 1 To d1Count
                    If d1Codes(j) = dVal Then
                        d1Counts(j) = d1Counts(j) + 1
                        found = True
                        Exit For
                    End If
                Next j
                If Not found Then
                    d1Count = d1Count + 1
                    ReDim Preserve d1Codes(1 To d1Count)
                    ReDim Preserve d1Counts(1 To d1Count)
                    d1Codes(d1Count) = dVal
                    d1Counts(d1Count) = 1
                End If
            End If
        End If
    Next i
    If d1Count > 0 Then
        For i = 1 To d1Count
            ws1.Cells(3 + i, 1).Value = d1Codes(i)
            ws1.Cells(3 + i, 2).Value = d1Counts(i)
        Next i
        ws1.Columns("A:B").AutoFit
        Dim r1 As Range
        Set r1 = ws1.Range(ws1.Cells(3, 1), ws1.Cells(3 + d1Count, 2))
        Dim c1 As ChartObject
        Set c1 = ws1.ChartObjects.Add(Left:=200, Top:=30, Width:=420, Height:=300)
        With c1.Chart
            .SetSourceData Source:=r1
            .ChartType = xlPie
            .HasTitle = True
            .ChartTitle.Text = "Supplier Anomaly Count Per Disposition"
            With .ChartTitle.Font
                .Size = 13
                .Bold = True
                .Color = RGB(31, 73, 125)
            End With
            .SeriesCollection(1).Name = selSup
            .SeriesCollection(1).HasDataLabels = True
            With .SeriesCollection(1).DataLabels
                .ShowCategoryName = True
                .ShowPercentage = True
                .ShowValue = False
                .Separator = Chr(10)
                .Font.Size = 9
            End With
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            .Legend.Font.Size = 9
            .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
            .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        End With
    Else
        ws1.Cells(4, 1).Value = "No data found for selected supplier."
    End If
    Application.DisplayAlerts = False
    For Each tempWs In ThisWorkbook.Worksheets
        If tempWs.Name = s2Name Then
            tempWs.Delete
            Exit For
        End If
    Next tempWs
    Application.DisplayAlerts = True
    Set ws2 = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error Resume Next
    ws2.Name = s2Name
    If ws2.Name <> s2Name Then ws2.Name = Left(s2Name, 31)
    On Error GoTo 0
    With ws2.Cells(1, 1)
        .Value = "Supplier:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    For i = 1 To supplierCount
        ws2.Cells(i, 26).Value = suppliers(i)
    Next i
    ws2.Cells(1, 2).Value = suppliers(1)
    With ws2.Cells(1, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=$Z$1:$Z$" & supplierCount
        .InCellDropdown = True
    End With
    With ws2.Cells(1, 2)
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End With
    With ws2.Cells(3, 1)
        .Value = "Defect Code"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    With ws2.Cells(3, 2)
        .Value = "Count"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws2.Columns("A:B").AutoFit
    ws2.Columns("Z").ColumnWidth = 0.1
    Dim btn2 As Button
    Set btn2 = ws2.Buttons.Add(ws2.Cells(1, 3).Left, ws2.Cells(1, 3).Top, 100, 22)
    btn2.Caption = "Refresh"
    btn2.OnAction = "RefreshSheet2"
    Dim d2Codes() As String
    Dim d2Counts() As Long
    Dim d2Count As Integer
    d2Count = 0
    For i = 4 To lastRow
        If SafeText(dataWs.Cells(i, "A")) = selSup Then
            dVal = SafeText(dataWs.Cells(i, 23))
            If dVal <> "" Then
                found = False
                For j = 1 To d2Count
                    If d2Codes(j) = dVal Then
                        d2Counts(j) = d2Counts(j) + 1
                        found = True
                        Exit For
                    End If
                Next j
                If Not found Then
                    d2Count = d2Count + 1
                    ReDim Preserve d2Codes(1 To d2Count)
                    ReDim Preserve d2Counts(1 To d2Count)
                    d2Codes(d2Count) = dVal
                    d2Counts(d2Count) = 1
                End If
            End If
        End If
    Next i
    If d2Count > 0 Then
        For i = 1 To d2Count
            ws2.Cells(3 + i, 1).Value = d2Codes(i)
            ws2.Cells(3 + i, 2).Value = d2Counts(i)
        Next i
        ws2.Columns("A:B").AutoFit
        Dim r2 As Range
        Set r2 = ws2.Range(ws2.Cells(3, 1), ws2.Cells(3 + d2Count, 2))
        Dim c2 As ChartObject
        Set c2 = ws2.ChartObjects.Add(Left:=200, Top:=30, Width:=420, Height:=300)
        With c2.Chart
            .SetSourceData Source:=r2
            .ChartType = xlPie
            .HasTitle = True
            .ChartTitle.Text = "Supplier Anomaly Count Per Defect Code"
            With .ChartTitle.Font
                .Size = 13
                .Bold = True
                .Color = RGB(31, 73, 125)
            End With
            .SeriesCollection(1).Name = selSup
            .SeriesCollection(1).HasDataLabels = True
            With .SeriesCollection(1).DataLabels
                .ShowCategoryName = True
                .ShowPercentage = True
                .ShowValue = False
                .Separator = Chr(10)
                .Font.Size = 9
            End With
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            .Legend.Font.Size = 9
            .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
            .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        End With
    Else
        ws2.Cells(4, 1).Value = "No data found for selected supplier."
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Done! Cleanup and both analysis sheets completed.", vbInformation, "DataCleanup"
End Sub
Function SafeText(c As Range) As String
    On Error Resume Next
    If IsError(c.Value) Then
        SafeText = ""
    Else
        SafeText = Trim(CStr(c.Value))
    End If
    On Error GoTo 0
End Function
Sub RefreshSheet1()
    Dim dataWs As Worksheet
    Dim analysisWs As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String
    Dim otherName As String
    Dim selSup As String
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    Dim dVal As String
    Dim found As Boolean
    sheetName = "Anomaly Count Per Disposition"
    otherName = "Defect Code Analysis"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then Set analysisWs = ws
        If ws.Name <> sheetName And ws.Name <> otherName Then
            If dataWs Is Nothing Then Set dataWs = ws
        End If
    Next ws
    If analysisWs Is Nothing Or dataWs Is Nothing Then
        MsgBox "Could not find required sheets. Please run DataCleanup first.", vbExclamation
        Exit Sub
    End If
    selSup = SafeText(analysisWs.Cells(1, 2))
    If selSup = "" Then
        MsgBox "Please select a supplier from the dropdown in B1.", vbExclamation
        Exit Sub
    End If
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    Dim dCodes() As String
    Dim dCounts() As Long
    Dim dCount As Integer
    dCount = 0
    For i = 4 To lastRow
        If SafeText(dataWs.Cells(i, "A")) = selSup Then
            dVal = SafeText(dataWs.Cells(i, 24))
            If dVal <> "" Then
                found = False
                For j = 1 To dCount
                    If dCodes(j) = dVal Then
                        dCounts(j) = dCounts(j) + 1
                        found = True
                        Exit For
                    End If
                Next j
                If Not found Then
                    dCount = dCount + 1
                    ReDim Preserve dCodes(1 To dCount)
                    ReDim Preserve dCounts(1 To dCount)
                    dCodes(dCount) = dVal
                    dCounts(dCount) = 1
                End If
            End If
        End If
    Next i
    analysisWs.Range("A4:B1000").ClearContents
    Dim co As ChartObject
    For Each co In analysisWs.ChartObjects
        co.Delete
    Next co
    If dCount = 0 Then
        analysisWs.Cells(4, 1).Value = "No data found for selected supplier."
        Exit Sub
    End If
    For i = 1 To dCount
        analysisWs.Cells(3 + i, 1).Value = dCodes(i)
        analysisWs.Cells(3 + i, 2).Value = dCounts(i)
    Next i
    analysisWs.Columns("A:B").AutoFit
    Dim r1 As Range
    Set r1 = analysisWs.Range(analysisWs.Cells(3, 1), analysisWs.Cells(3 + dCount, 2))
    Dim c1 As ChartObject
    Set c1 = analysisWs.ChartObjects.Add(Left:=200, Top:=30, Width:=420, Height:=300)
    With c1.Chart
        .SetSourceData Source:=r1
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Supplier Anomaly Count Per Disposition"
        With .ChartTitle.Font
            .Size = 13
            .Bold = True
            .Color = RGB(31, 73, 125)
        End With
        .SeriesCollection(1).Name = selSup
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowCategoryName = True
            .ShowPercentage = True
            .ShowValue = False
            .Separator = Chr(10)
            .Font.Size = 9
        End With
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub
Sub RefreshSheet2()
    Dim dataWs As Worksheet
    Dim analysisWs As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String
    Dim otherName As String
    Dim selSup As String
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    Dim dVal As String
    Dim found As Boolean
    sheetName = "Defect Code Analysis"
    otherName = "Anomaly Count Per Disposition"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then Set analysisWs = ws
        If ws.Name <> sheetName And ws.Name <> otherName Then
            If dataWs Is Nothing Then Set dataWs = ws
        End If
    Next ws
    If analysisWs Is Nothing Or dataWs Is Nothing Then
        MsgBox "Could not find required sheets. Please run DataCleanup first.", vbExclamation
        Exit Sub
    End If
    selSup = SafeText(analysisWs.Cells(1, 2))
    If selSup = "" Then
        MsgBox "Please select a supplier from the dropdown in B1.", vbExclamation
        Exit Sub
    End If
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    Dim dCodes() As String
    Dim dCounts() As Long
    Dim dCount As Integer
    dCount = 0
    For i = 4 To lastRow
        If SafeText(dataWs.Cells(i, "A")) = selSup Then
            dVal = SafeText(dataWs.Cells(i, 23))
            If dVal <> "" Then
                found = False
                For j = 1 To dCount
                    If dCodes(j) = dVal Then
                        dCounts(j) = dCounts(j) + 1
                        found = True
                        Exit For
                    End If
                Next j
                If Not found Then
                    dCount = dCount + 1
                    ReDim Preserve dCodes(1 To dCount)
                    ReDim Preserve dCounts(1 To dCount)
                    dCodes(dCount) = dVal
                    dCounts(dCount) = 1
                End If
            End If
        End If
    Next i
    analysisWs.Range("A4:B1000").ClearContents
    Dim co As ChartObject
    For Each co In analysisWs.ChartObjects
        co.Delete
    Next co
    If dCount = 0 Then
        analysisWs.Cells(4, 1).Value = "No data found for selected supplier."
        Exit Sub
    End If
    For i = 1 To dCount
        analysisWs.Cells(3 + i, 1).Value = dCodes(i)
        analysisWs.Cells(3 + i, 2).Value = dCounts(i)
    Next i
    analysisWs.Columns("A:B").AutoFit
    Dim r2 As Range
    Set r2 = analysisWs.Range(analysisWs.Cells(3, 1), analysisWs.Cells(3 + dCount, 2))
    Dim c2 As ChartObject
    Set c2 = analysisWs.ChartObjects.Add(Left:=200, Top:=30, Width:=420, Height:=300)
    With c2.Chart
        .SetSourceData Source:=r2
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Supplier Anomaly Count Per Defect Code"
        With .ChartTitle.Font
            .Size = 13
            .Bold = True
            .Color = RGB(31, 73, 125)
        End With
        .SeriesCollection(1).Name = selSup
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowCategoryName = True
            .ShowPercentage = True
            .ShowValue = False
            .Separator = Chr(10)
            .Font.Size = 9
        End With
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub
