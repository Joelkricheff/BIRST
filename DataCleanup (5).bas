Sub DataCleanup()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    For i = lastRow To 4 Step -1
        If Trim(ws.Cells(i, "H").Value) <> "CLOSED" Then
            ws.Rows(i).Delete
        End If
    Next i

       lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    For i = lastRow To 4 Step -1
        If Left(Trim(CStr(ws.Cells(i, "U").Value)), 1) <> "4" Then
            ws.Rows(i).Delete
        End If
    Next i

       ws.Columns(23).Insert Shift:=xlToRight

      ws.Cells(3, 23).Value = "Extracted"

    lastRow = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row
    If lastRow >= 4 Then
        ws.Range(ws.Cells(4, 23), ws.Cells(lastRow, 23)).Formula = _
            "=LEFT(V4,FIND(""-"",V4&""-"")-1)"
    End If

     lastRow = ws.Cells(ws.Rows.Count, 23).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(Left(Trim(ws.Cells(i, 23).Value), 9)) = "NO DEFECT" Then
            ws.Rows(i).Delete
        End If
    Next i

      lastRow = ws.Cells(ws.Rows.Count, 24).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(Trim(ws.Cells(i, 24).Value)) = "PROMOTE TO HAR" Then
            ws.Rows(i).Delete
        End If
    Next i

     lastRow = ws.Cells(ws.Rows.Count, 24).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(Trim(ws.Cells(i, 24).Value)) = "NDF" Then
            ws.Rows(i).Delete
        End If
    Next i

     Call BuildAnalysisSheet(ws)

       Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done! All steps completed and analysis sheet created.", vbInformation, "DataCleanup"

End Sub


Sub BuildAnalysisSheet(dataWs As Worksheet)

    Dim analysisWs As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer

    sheetName = "Supplier Anomaly Analysis"  ' <-- change sheet name here only

    ' Delete existing analysis sheet if present
    Application.DisplayAlerts = False
    Dim tempWs As Worksheet
    For Each tempWs In ThisWorkbook.Worksheets
        If tempWs.Name = sheetName Then
            tempWs.Delete
            Exit For
        End If
    Next tempWs
    Application.DisplayAlerts = True

    ' Create new sheet and name it
    Set analysisWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    analysisWs.Name = sheetName
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row

    Dim suppliers() As String
    Dim supplierCount As Integer
    supplierCount = 0

    Dim cellVal As String
    Dim found As Boolean

    For i = 4 To lastRow
        cellVal = Trim(dataWs.Cells(i, "A").Value)
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
        MsgBox "No supplier data found in column A. Check that cleanup ran correctly.", vbExclamation
        Exit Sub
    End If

      With analysisWs.Cells(1, 1)
        .Value = "Supplier:"
        .Font.Bold = True
        .Font.Size = 12
    End With

    ' Write supplier list to helper column Z for validation source
    For i = 1 To supplierCount
        analysisWs.Cells(i, 26).Value = suppliers(i)
    Next i

    ' Default B1 to first supplier
    analysisWs.Cells(1, 2).Value = suppliers(1)

    ' Apply dropdown validation to B1
    With analysisWs.Cells(1, 2).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=$Z$1:$Z$" & supplierCount
        .InCellDropdown = True
    End With

    With analysisWs.Cells(1, 2)
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End With
    With analysisWs.Cells(3, 1)
        .Value = "Defect Code"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    With analysisWs.Cells(3, 2)
        .Value = "Count"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With

    analysisWs.Columns("A:B").AutoFit
    analysisWs.Columns("Z").ColumnWidth = 0.1
    Dim btn As Button
    Set btn = analysisWs.Buttons.Add( _
        analysisWs.Cells(1, 3).Left, _
        analysisWs.Cells(1, 3).Top, 100, 22)
    btn.Caption = "Refresh"
    btn.OnAction = "RefreshAnalysis"

       Call PopulateAnalysis(dataWs, analysisWs)

End Sub


Sub RefreshAnalysis()

    Dim sheetName As String
    Dim analysisWs As Worksheet
    Dim dataWs As Worksheet
    Dim ws As Worksheet

    sheetName = "Supplier Anomaly Analysis"  ' <-- must match name above

    ' Find analysis sheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set analysisWs = ws
        End If
    Next ws

    If analysisWs Is Nothing Then
        MsgBox "Analysis sheet not found. Please run DataCleanup first.", vbExclamation
        Exit Sub
    End If

    ' Find data sheet (first sheet that is not the analysis sheet)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetName Then
            Set dataWs = ws
            Exit For
        End If
    Next ws

    Call PopulateAnalysis(dataWs, analysisWs)

End Sub

Sub PopulateAnalysis(dataWs As Worksheet, analysisWs As Worksheet)

    Dim selectedSupplier As String
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer

    selectedSupplier = Trim(analysisWs.Cells(1, 2).Value)

    If selectedSupplier = "" Then
        MsgBox "Please select a supplier from the dropdown in cell B1.", vbExclamation
        Exit Sub
    End If

       lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row

    Dim defectCodes() As String
    Dim defectCounts() As Long
    Dim defectCount As Integer
    defectCount = 0

    Dim supplierVal As String
    Dim defectVal As String
    Dim found As Boolean

    For i = 4 To lastRow
        supplierVal = Trim(dataWs.Cells(i, "A").Value)
        defectVal = Trim(dataWs.Cells(i, 24).Value)

        If supplierVal = selectedSupplier And defectVal <> "" Then
            found = False
            For j = 1 To defectCount
                If defectCodes(j) = defectVal Then
                    defectCounts(j) = defectCounts(j) + 1
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                defectCount = defectCount + 1
                ReDim Preserve defectCodes(1 To defectCount)
                ReDim Preserve defectCounts(1 To defectCount)
                defectCodes(defectCount) = defectVal
                defectCounts(defectCount) = 1
            End If
        End If
    Next i

        analysisWs.Range("A4:B1000").ClearContents

    ' Remove existing chart
    Dim chtObj As ChartObject
    For Each chtObj In analysisWs.ChartObjects
        chtObj.Delete
    Next chtObj

    If defectCount = 0 Then
        analysisWs.Cells(4, 1).Value = "No data found for selected supplier."
        Exit Sub
    End If

        For i = 1 To defectCount
        analysisWs.Cells(3 + i, 1).Value = defectCodes(i)
        analysisWs.Cells(3 + i, 2).Value = defectCounts(i)
    Next i

    analysisWs.Columns("A:B").AutoFit

        Dim pivotRange As Range
    Set pivotRange = analysisWs.Range( _
        analysisWs.Cells(3, 1), _
        analysisWs.Cells(3 + defectCount, 2))

    Dim co As ChartObject
    Set co = analysisWs.ChartObjects.Add(Left:=200, Top:=30, Width:=420, Height:=300)

    With co.Chart
        .SetSourceData Source:=pivotRange
        .ChartType = xlPie

        .HasTitle = True
        .ChartTitle.Text = "Supplier Anomaly Count Per Defect Code"
        With .ChartTitle.Font
            .Size = 13
            .Bold = True
            .Color = RGB(31, 73, 125)
        End With

        .SeriesCollection(1).Name = selectedSupplier

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
