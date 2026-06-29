Public Const S1 As String = "Anomaly Count Per Disposition"
Public Const S2 As String = "Defect Code Analysis"
Public Const S3 As String = "Anomaly Count Per Project"
Public Const KEYSHEET As String = "Project Key"
Public Const PROJNUMCOL As Long = 35
Sub DataCleanup()
    Dim dataWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Integer
    Dim cellVal As String
    Dim found As Boolean
    Set dataWs = ActiveSheet
    If IsHelperSheet(dataWs.Name) Then Set dataWs = GetDataSheet()
    If dataWs Is Nothing Then
        MsgBox "Could not locate the data sheet.", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    lastRow = dataWs.Cells(dataWs.Rows.Count, "H").End(xlUp).Row
    For i = lastRow To 4 Step -1
        If SafeText(dataWs.Cells(i, "H")) <> "CLOSED" Then dataWs.Rows(i).Delete
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, "U").End(xlUp).Row
    For i = lastRow To 4 Step -1
        If Left(SafeText(dataWs.Cells(i, "U")), 1) <> "4" Then dataWs.Rows(i).Delete
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
        If UCase(Left(SafeText(dataWs.Cells(i, 23)), 9)) = "NO DEFECT" Then dataWs.Rows(i).Delete
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, 24).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(SafeText(dataWs.Cells(i, 24))) = "PROMOTE TO HAR" Then dataWs.Rows(i).Delete
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, 24).End(xlUp).Row
    For i = lastRow To 4 Step -1
        If UCase(SafeText(dataWs.Cells(i, 24))) = "NDF" Then dataWs.Rows(i).Delete
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    Dim r3last As Long
    Dim r4last As Long
    Dim projNameCol As Long
    Dim cl As String
    r3last = dataWs.Cells(3, dataWs.Columns.Count).End(xlToLeft).Column
    r4last = dataWs.Cells(4, dataWs.Columns.Count).End(xlToLeft).Column
    projNameCol = r3last
    If r4last > projNameCol Then projNameCol = r4last
    projNameCol = projNameCol + 1
    dataWs.Cells(3, projNameCol).Value = "Project Name"
    cl = ColLetter(PROJNUMCOL)
    If lastRow >= 4 Then
        dataWs.Range(dataWs.Cells(4, projNameCol), dataWs.Cells(lastRow, projNameCol)).Formula = _
            "=IFERROR(VLOOKUP(LEFT(AI4,5),'Project Key'!A:B,2,FALSE),IFERROR(VLOOKUP(VALUE(LEFT(AI4,5)),'Project Key'!A:B,2,FALSE),""""))"
    End If
    Dim projShortCol As Long
    Dim pcl As String
    projShortCol = projNameCol + 1
    dataWs.Cells(3, projShortCol).Value = "Project Short"
    pcl = ColLetter(projNameCol)
    If lastRow >= 4 Then
        dataWs.Range(dataWs.Cells(4, projShortCol), dataWs.Cells(lastRow, projShortCol)).Formula = _
            "=IFERROR(LEFT(" & pcl & "4,FIND("" "",SUBSTITUTE(" & pcl & "4,""- "",""-""&CHAR(1)))-1)," & pcl & "4)"
    End If
    Application.Calculate
    Dim suppliers() As String
    Dim supplierCount As Integer
    supplierCount = 0
    For i = 4 To lastRow
        cellVal = SafeText(dataWs.Cells(i, "A"))
        If cellVal <> "" Then
            found = False
            For j = 1 To supplierCount
                If suppliers(j) = cellVal Then found = True: Exit For
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
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Set ws1 = BuildScaffold(dataWs, S1, "Disposition Code", "RefreshSheet1", suppliers, supplierCount)
    PopulateChart dataWs, ws1, 24, "Supplier Anomaly Count Per Disposition", suppliers(1), lastRow
    Set ws2 = BuildScaffold(dataWs, S2, "Defect Code", "RefreshSheet2", suppliers, supplierCount)
    PopulateChart dataWs, ws2, 23, "Supplier Anomaly Count Per Defect Code", suppliers(1), lastRow
    Set ws3 = BuildScaffold(dataWs, S3, "Project Name", "RefreshSheet3", suppliers, supplierCount)
    PopulateChart dataWs, ws3, projNameCol, "Anomaly Count Per Project", suppliers(1), lastRow
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Done! Cleanup, project lookup, and all three analysis sheets completed.", vbInformation, "DataCleanup"
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
Function SafeVar(v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        SafeVar = ""
    ElseIf IsNull(v) Then
        SafeVar = ""
    Else
        SafeVar = Trim(CStr(v))
    End If
    On Error GoTo 0
End Function
Function ColLetter(n As Long) As String
    Dim s As String
    Dim r As Long
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(65 + r) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function
Function IsHelperSheet(nm As String) As Boolean
    IsHelperSheet = (nm = S1 Or nm = S2 Or nm = S3 Or nm = KEYSHEET)
End Function
Function GetDataSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsHelperSheet(ws.Name) Then Set GetDataSheet = ws: Exit Function
    Next ws
End Function
Function GetSheet(nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function
Function FindColByHeader(dataWs As Worksheet, hdr As String) As Long
    Dim c As Long
    Dim lastC As Long
    lastC = dataWs.Cells(3, dataWs.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If SafeText(dataWs.Cells(3, c)) = hdr Then FindColByHeader = c: Exit Function
    Next c
    FindColByHeader = 0
End Function
Function BuildScaffold(dataWs As Worksheet, sheetName As String, headerLabel As String, refreshMacro As String, suppliers() As String, supplierCount As Integer) As Worksheet
    Dim wsx As Worksheet
    Dim t As Worksheet
    Dim i As Long
    Application.DisplayAlerts = False
    For Each t In ThisWorkbook.Worksheets
        If t.Name = sheetName Then t.Delete: Exit For
    Next t
    Application.DisplayAlerts = True
    Set wsx = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    On Error Resume Next
    wsx.Name = sheetName
    If wsx.Name <> sheetName Then wsx.Name = Left(sheetName, 31)
    On Error GoTo 0
    With wsx.Cells(1, 1)
        .Value = "Supplier:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    For i = 1 To supplierCount
        wsx.Cells(i, 26).Value = suppliers(i)
    Next i
    wsx.Cells(1, 2).Value = suppliers(1)
    With wsx.Cells(1, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=$Z$1:$Z$" & supplierCount
        .InCellDropdown = True
    End With
    With wsx.Cells(1, 2)
        .Font.Size = 12
        .Interior.Color = RGB(220, 230, 241)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End With
    With wsx.Cells(3, 1)
        .Value = headerLabel
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    With wsx.Cells(3, 2)
        .Value = "Count"
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsx.Columns("A:B").AutoFit
    wsx.Columns("Z").ColumnWidth = 0.1
    Dim b As Button
    Set b = wsx.Buttons.Add(wsx.Cells(1, 3).Left, wsx.Cells(1, 3).Top, 100, 22)
    b.Caption = "Refresh"
    b.OnAction = refreshMacro
    Set BuildScaffold = wsx
End Function
Sub PopulateChart(dataWs As Worksheet, analysisWs As Worksheet, countCol As Long, chartTitle As String, selSup As String, lastRow As Long)
    Dim i As Long
    Dim r As Long
    Dim n As Long
    Dim nRows As Long
    Dim pos As Long
    Dim sA As String
    Dim sC As String
    Dim codes() As String
    Dim counts() As Long
    Dim arrA As Variant
    Dim arrC As Variant
    Dim mapC As New Collection
    n = 0
    If countCol >= 1 And lastRow >= 4 Then
        arrA = dataWs.Range(dataWs.Cells(4, 1), dataWs.Cells(lastRow, 1)).Value
        arrC = dataWs.Range(dataWs.Cells(4, countCol), dataWs.Cells(lastRow, countCol)).Value
        If IsArray(arrA) Then nRows = UBound(arrA, 1) Else nRows = 1
        For r = 1 To nRows
            If IsArray(arrA) Then
                sA = SafeVar(arrA(r, 1))
                sC = SafeVar(arrC(r, 1))
            Else
                sA = SafeVar(arrA)
                sC = SafeVar(arrC)
            End If
            If sA = selSup And sC <> "" Then
                pos = 0
                On Error Resume Next
                pos = mapC(sC)
                On Error GoTo 0
                If pos = 0 Then
                    n = n + 1
                    ReDim Preserve codes(1 To n)
                    ReDim Preserve counts(1 To n)
                    codes(n) = sC
                    counts(n) = 1
                    mapC.Add n, sC
                Else
                    counts(pos) = counts(pos) + 1
                End If
            End If
        Next r
    End If
    analysisWs.Range("A4:B100000").ClearContents
    Dim co As ChartObject
    For Each co In analysisWs.ChartObjects
        co.Delete
    Next co
    If n = 0 Then
        analysisWs.Cells(4, 1).Value = "No data found for selected supplier."
        Exit Sub
    End If
    For i = 1 To n
        analysisWs.Cells(3 + i, 1).Value = codes(i)
        analysisWs.Cells(3 + i, 2).Value = counts(i)
    Next i
    analysisWs.Columns("A:B").AutoFit
    Dim rng As Range
    Set rng = analysisWs.Range(analysisWs.Cells(3, 1), analysisWs.Cells(3 + n, 2))
    Dim ch As ChartObject
    Set ch = analysisWs.ChartObjects.Add(Left:=200, Top:=30, Width:=420, Height:=300)
    With ch.Chart
        .SetSourceData Source:=rng
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .ChartTitle.Font.Size = 13
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(31, 73, 125)
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
Sub RefreshSheet1()
    Dim dataWs As Worksheet
    Dim aws As Worksheet
    Dim selSup As String
    Dim lastRow As Long
    Set dataWs = GetDataSheet()
    Set aws = GetSheet(S1)
    If dataWs Is Nothing Or aws Is Nothing Then MsgBox "Run DataCleanup first.", vbExclamation: Exit Sub
    selSup = SafeText(aws.Cells(1, 2))
    If selSup = "" Then MsgBox "Select a supplier in B1.", vbExclamation: Exit Sub
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    PopulateChart dataWs, aws, 24, "Supplier Anomaly Count Per Disposition", selSup, lastRow
End Sub
Sub RefreshSheet2()
    Dim dataWs As Worksheet
    Dim aws As Worksheet
    Dim selSup As String
    Dim lastRow As Long
    Set dataWs = GetDataSheet()
    Set aws = GetSheet(S2)
    If dataWs Is Nothing Or aws Is Nothing Then MsgBox "Run DataCleanup first.", vbExclamation: Exit Sub
    selSup = SafeText(aws.Cells(1, 2))
    If selSup = "" Then MsgBox "Select a supplier in B1.", vbExclamation: Exit Sub
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    PopulateChart dataWs, aws, 23, "Supplier Anomaly Count Per Defect Code", selSup, lastRow
End Sub
Sub RefreshSheet3()
    Dim dataWs As Worksheet
    Dim aws As Worksheet
    Dim selSup As String
    Dim lastRow As Long
    Dim projCol As Long
    Set dataWs = GetDataSheet()
    Set aws = GetSheet(S3)
    If dataWs Is Nothing Or aws Is Nothing Then MsgBox "Run DataCleanup first.", vbExclamation: Exit Sub
    selSup = SafeText(aws.Cells(1, 2))
    If selSup = "" Then MsgBox "Select a supplier in B1.", vbExclamation: Exit Sub
    projCol = FindColByHeader(dataWs, "Project Name")
    If projCol = 0 Then MsgBox "Project Name column not found. Run DataCleanup first.", vbExclamation: Exit Sub
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    PopulateChart dataWs, aws, projCol, "Anomaly Count Per Project", selSup, lastRow
End Sub
