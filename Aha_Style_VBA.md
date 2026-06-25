Aha Style Output VBA:

Option Explicit

Sub BuildOperationsRoadmap_AhaStyle_SeparateWorkbook_WithRefresh()

    ' ==========================================================
    ' PURPOSE
    ' Refreshes Power Query / Aha REST API data first,
    ' then creates an Aha-style roadmap in a BRAND-NEW workbook.
    '
    ' IMPORTANT:
    ' - This macro reads from the current workbook.
    ' - This macro does NOT delete or overwrite any current sheets.
    ' - Output is created in a separate new workbook.
    '
    ' Expected source sheet names:
    ' 1. Tabled Data (M Query)
    ' 2. Query1
    ' 3. Sheet1
    '
    ' Expected source columns:
    ' - MFS Category
    ' - Supporting Objectives
    ' - Initiative Alignment
    ' - Initiative name
    ' - Quarter
    ' - Status
    ' Optional:
    ' - Priority
    ' - OPS Rank
    ' ==========================================================

    Dim wbSrc As Workbook
    Dim wbOut As Workbook
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet

    Dim colMap As Object
    Dim dict As Object
    Dim rec As Object
    Dim qDict As Object
    Dim qOrder As Object

    Dim lastCol As Long
    Dim lastRow As Long
    Dim data As Variant

    Dim colCategory As Long
    Dim colSupporting As Long
    Dim colAlignment As Long
    Dim colName As Long
    Dim colQuarter As Long
    Dim colStatus As Long
    Dim colPriority As Long
    Dim colOpsRank As Long

    Dim hasPriority As Boolean
    Dim hasOpsRank As Boolean

    Dim quarters As Variant
    Dim keys As Variant

    Dim i As Long
    Dim c As Long
    Dim r As Long
    Dim qIndex As Long

    Dim cat As String
    Dim supp As String
    Dim align As String
    Dim nameVal As String
    Dim qtrName As String
    Dim statusVal As String
    Dim priorityVal As String
    Dim opsRankVal As String
    Dim key As String

    Dim k As Variant
    Dim qCol As Variant

    Dim startQ As String
    Dim endQ As String
    Dim startIdx As Long
    Dim endIdx As Long

    Dim leftPos As Double
    Dim topPos As Double
    Dim barTop As Double
    Dim labelTop As Double
    Dim barWidth As Double
    Dim barHeight As Double
    Dim labelWidth As Double

    Dim barShape As Shape
    Dim labelShape As Shape

    Dim timelineStartCol As Long
    Dim timelineEndCol As Long

    Dim calcMode As XlCalculation

    On Error GoTo CleanFail

    ' ==========================================================
    ' PRESERVE CURRENT APP SETTINGS
    ' ==========================================================
    Set wbSrc = ThisWorkbook

    calcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Refreshing Aha / Power Query data..."

    ' ==========================================================
    ' AUTO REFRESH POWER QUERY / AHA API DATA
    ' ==========================================================
    ' This refreshes all workbook connections, including Power Query
    ' queries that pull from the Aha REST API.
    wbSrc.RefreshAll

    ' Wait for async Power Query refreshes to complete.
    DoEvents
    Application.CalculateUntilAsyncQueriesDone
    DoEvents

    Application.StatusBar = "Aha / Power Query refresh complete. Building roadmap..."

    ' ==========================================================
    ' SOURCE SHEET DETECTION
    ' ==========================================================
    Set wsSrc = Nothing

    If SheetExists(wbSrc, "Tabled Data (M Query)") Then
        Set wsSrc = wbSrc.Worksheets("Tabled Data (M Query)")
    ElseIf SheetExists(wbSrc, "Query1") Then
        Set wsSrc = wbSrc.Worksheets("Query1")
    ElseIf SheetExists(wbSrc, "Sheet1") Then
        Set wsSrc = wbSrc.Worksheets("Sheet1")
    Else
        Err.Raise vbObjectError + 1000, , _
            "Could not find a source sheet named 'Tabled Data (M Query)', 'Query1', or 'Sheet1'."
    End If

    ' ==========================================================
    ' DEFINE QUARTERS
    ' ==========================================================
    quarters = Array( _
        "Q.FY26 Q1", "Q.FY26 Q2", "Q.FY26 Q3", "Q.FY26 Q4", _
        "Q.FY27 Q1", "Q.FY27 Q2", "Q.FY27 Q3", "Q.FY27 Q4", _
        "Q.FY28 Q1", "Q.FY28 Q2", "Q.FY28 Q3", "Q.FY28 Q4")

    Set qOrder = CreateObject("Scripting.Dictionary")

    For i = LBound(quarters) To UBound(quarters)
        qOrder(CStr(quarters(i))) = i + 1
    Next i

    ' ==========================================================
    ' MAP SOURCE COLUMNS
    ' ==========================================================
    Set colMap = CreateObject("Scripting.Dictionary")

    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Len(Trim(CStr(wsSrc.Cells(1, c).Value))) > 0 Then
            colMap(Trim(CStr(wsSrc.Cells(1, c).Value))) = c
        End If
    Next c

    If Not colMap.exists("MFS Category") Then Err.Raise vbObjectError + 1001, , "Missing column: MFS Category"
    If Not colMap.exists("Supporting Objectives") Then Err.Raise vbObjectError + 1002, , "Missing column: Supporting Objectives"
    If Not colMap.exists("Initiative Alignment") Then Err.Raise vbObjectError + 1003, , "Missing column: Initiative Alignment"
    If Not colMap.exists("Initiative name") Then Err.Raise vbObjectError + 1004, , "Missing column: Initiative name"
    If Not colMap.exists("Quarter") Then Err.Raise vbObjectError + 1005, , "Missing column: Quarter"
    If Not colMap.exists("Status") Then Err.Raise vbObjectError + 1006, , "Missing column: Status"

    colCategory = colMap("MFS Category")
    colSupporting = colMap("Supporting Objectives")
    colAlignment = colMap("Initiative Alignment")
    colName = colMap("Initiative name")
    colQuarter = colMap("Quarter")
    colStatus = colMap("Status")

    hasPriority = colMap.exists("Priority")
    hasOpsRank = colMap.exists("OPS Rank")

    If hasPriority Then colPriority = colMap("Priority")
    If hasOpsRank Then colOpsRank = colMap("OPS Rank")

    ' ==========================================================
    ' LOAD SOURCE DATA AFTER REFRESH
    ' ==========================================================
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        Err.Raise vbObjectError + 1007, , "The source sheet has no data rows after refresh."
    End If

    data = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Value

    ' ==========================================================
    ' BUILD MODEL
    ' One record per initiative.
    ' This identifies the earliest quarter and latest quarter
    ' where the initiative has C/D/P status.
    ' ==========================================================
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(data, 1)

        cat = Trim(CStr(data(i, colCategory)))
        supp = Trim(CStr(data(i, colSupporting)))
        align = Trim(CStr(data(i, colAlignment)))
        nameVal = Trim(CStr(data(i, colName)))
        qtrName = Trim(CStr(data(i, colQuarter)))
        statusVal = UCase(Trim(CStr(data(i, colStatus))))

        If hasPriority Then
            priorityVal = Trim(CStr(data(i, colPriority)))
        Else
            priorityVal = ""
        End If

        If hasOpsRank Then
            opsRankVal = Trim(CStr(data(i, colOpsRank)))
        Else
            opsRankVal = ""
        End If

        If Len(nameVal) = 0 Then GoTo SkipRow
        If statusVal <> "C" And statusVal <> "D" And statusVal <> "P" Then GoTo SkipRow
        If Not qOrder.exists(qtrName) Then GoTo SkipRow

        key = cat & "|" & supp & "|" & align & "|" & nameVal

        If Not dict.exists(key) Then

            Set rec = CreateObject("Scripting.Dictionary")

            rec.Add "MFS Category", cat
            rec.Add "Supporting Initiative", supp
            rec.Add "Initiative Alignment", align
            rec.Add "Initiative Name", nameVal
            rec.Add "Priority", priorityVal
            rec.Add "OPS Rank", opsRankVal
            rec.Add "StartIndex", 9999
            rec.Add "EndIndex", -1
            rec.Add "StartQuarter", ""
            rec.Add "EndQuarter", ""

            Set qDict = CreateObject("Scripting.Dictionary")

            For Each qCol In quarters
                qDict.Add CStr(qCol), ""
            Next qCol

            rec.Add "Quarters", qDict

            dict.Add key, rec

        End If

        Set rec = dict(key)
        Set qDict = rec("Quarters")

        qDict(qtrName) = statusVal

        If CLng(qOrder(qtrName)) < CLng(rec("StartIndex")) Then
            rec("StartIndex") = CLng(qOrder(qtrName))
            rec("StartQuarter") = qtrName
        End If

        If CLng(qOrder(qtrName)) > CLng(rec("EndIndex")) Then
            rec("EndIndex") = CLng(qOrder(qtrName))
            rec("EndQuarter") = qtrName
        End If

SkipRow:
    Next i

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 1008, , "No valid C, D, or P roadmap records were found after refresh."
    End If

    keys = dict.keys

    If dict.Count > 1 Then
        QuickSortVariant keys, LBound(keys), UBound(keys)
    End If

    ' ==========================================================
    ' CREATE BRAND-NEW OUTPUT WORKBOOK
    ' ==========================================================
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = "Aha_Style_Test_Output"

    timelineStartCol = 7
    timelineEndCol = timelineStartCol + UBound(quarters)

    ' ==========================================================
    ' HEADERS
    ' ==========================================================
    wsOut.Cells(1, 1).Value = "MFS Category"
    wsOut.Cells(1, 2).Value = "Supporting Initiative"
    wsOut.Cells(1, 3).Value = "Initiative Alignment"
    wsOut.Cells(1, 4).Value = "Initiative Name"
    wsOut.Cells(1, 5).Value = "Priority"
    wsOut.Cells(1, 6).Value = "OPS Rank"

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(2, 1)).Merge
    wsOut.Range(wsOut.Cells(1, 2), wsOut.Cells(2, 2)).Merge
    wsOut.Range(wsOut.Cells(1, 3), wsOut.Cells(2, 3)).Merge
    wsOut.Range(wsOut.Cells(1, 4), wsOut.Cells(2, 4)).Merge
    wsOut.Range(wsOut.Cells(1, 5), wsOut.Cells(2, 5)).Merge
    wsOut.Range(wsOut.Cells(1, 6), wsOut.Cells(2, 6)).Merge

    ' Year headers
    wsOut.Range(wsOut.Cells(1, timelineStartCol), wsOut.Cells(1, timelineStartCol + 3)).Merge
    wsOut.Cells(1, timelineStartCol).Value = "FY26"

    wsOut.Range(wsOut.Cells(1, timelineStartCol + 4), wsOut.Cells(1, timelineStartCol + 7)).Merge
    wsOut.Cells(1, timelineStartCol + 4).Value = "FY27"

    wsOut.Range(wsOut.Cells(1, timelineStartCol + 8), wsOut.Cells(1, timelineStartCol + 11)).Merge
    wsOut.Cells(1, timelineStartCol + 8).Value = "FY28"

    For qIndex = LBound(quarters) To UBound(quarters)
        wsOut.Cells(2, timelineStartCol + qIndex).Value = Replace(CStr(quarters(qIndex)), "Q.", "")
    Next qIndex

    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(2, timelineEndCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
    End With

    ' ==========================================================
    ' WRITE LEFT-SIDE DATA
    ' ==========================================================
    r = 3

    For Each k In keys

        Set rec = dict(k)

        wsOut.Cells(r, 1).Value = rec("MFS Category")
        wsOut.Cells(r, 2).Value = rec("Supporting Initiative")
        wsOut.Cells(r, 3).Value = rec("Initiative Alignment")
        wsOut.Cells(r, 4).Value = rec("Initiative Name")
        wsOut.Cells(r, 5).Value = rec("Priority")
        wsOut.Cells(r, 6).Value = rec("OPS Rank")

        wsOut.Rows(r).RowHeight = 38

        r = r + 1

    Next k

    ' ==========================================================
    ' FORMAT OUTPUT
    ' ==========================================================
    wsOut.Columns(1).ColumnWidth = 28
    wsOut.Columns(2).ColumnWidth = 24
    wsOut.Columns(3).ColumnWidth = 30
    wsOut.Columns(4).ColumnWidth = 36
    wsOut.Columns(5).ColumnWidth = 10
    wsOut.Columns(6).ColumnWidth = 10

    For c = timelineStartCol To timelineEndCol
        wsOut.Columns(c).ColumnWidth = 15
    Next c

    With wsOut.Range(wsOut.Cells(3, timelineStartCol), wsOut.Cells(r - 1, timelineEndCol))
        .Interior.Color = RGB(250, 250, 250)
        .Borders.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With

    With wsOut.Range(wsOut.Cells(3, 1), wsOut.Cells(r - 1, 6))
        .Borders.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' ==========================================================
    ' MERGE GROUPING COLUMNS IN NEW WORKBOOK ONLY
    ' ==========================================================
    MergeColumnSafe wsOut, 1, 3, r - 1
    MergeColumnSafe wsOut, 2, 3, r - 1
    MergeColumnSafe wsOut, 3, 3, r - 1

    ' ==========================================================
    ' DRAW AHA-STYLE LABELS AND BARS
    ' Label sits directly above the bar like Aha.
    ' ==========================================================
    r = 3

    For Each k In keys

        Set rec = dict(k)

        startQ = CStr(rec("StartQuarter"))
        endQ = CStr(rec("EndQuarter"))

        If Len(startQ) > 0 And Len(endQ) > 0 Then

            startIdx = CLng(rec("StartIndex"))
            endIdx = CLng(rec("EndIndex"))

            leftPos = wsOut.Cells(r, timelineStartCol + startIdx - 1).Left + 3
            topPos = wsOut.Rows(r).Top

            labelTop = topPos + 3
            barTop = topPos + 22
            barHeight = 10

            barWidth = _
                wsOut.Cells(r, timelineStartCol + endIdx - 1).Left + _
                wsOut.Cells(r, timelineStartCol + endIdx - 1).Width - _
                leftPos - 5

            If barWidth < 24 Then barWidth = 24

            labelWidth = barWidth

            ' Allows readable names even when a bar is only one quarter wide.
            If labelWidth < 190 Then labelWidth = 190

            ' ------------------------------
            ' Initiative label directly above bar
            ' ------------------------------
            Set labelShape = wsOut.Shapes.AddTextbox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=leftPos, _
                Top:=labelTop, _
                Width:=labelWidth, _
                Height:=16)

            With labelShape
                .Name = "AhaLabel_" & CStr(r)
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
                .Placement = xlMoveAndSize

                With .TextFrame2
                    .MarginLeft = 0
                    .MarginRight = 0
                    .MarginTop = 0
                    .MarginBottom = 0
                    .VerticalAnchor = msoAnchorMiddle
                    .TextRange.Text = CStr(rec("Initiative Name"))
                    .TextRange.Font.Size = 8
                    .TextRange.Font.Bold = msoTrue
                    .TextRange.Font.Fill.ForeColor.RGB = RGB(55, 55, 55)
                End With
            End With

            ' ------------------------------
            ' Timeline bar underneath label
            ' ------------------------------
            Set barShape = wsOut.Shapes.AddShape( _
                Type:=msoShapeRoundedRectangle, _
                Left:=leftPos, _
                Top:=barTop, _
                Width:=barWidth, _
                Height:=barHeight)

            With barShape
                .Name = "AhaBar_" & CStr(r)
                .Fill.ForeColor.RGB = GetStatusColor(GetPrimaryStatus(rec("Quarters"), startQ, quarters))
                .Fill.Transparency = 0.08
                .Line.ForeColor.RGB = RGB(120, 120, 120)
                .Line.Weight = 0.75
                .Placement = xlMoveAndSize
            End With

        End If

        r = r + 1

    Next k

    ' ==========================================================
    ' LEGEND
    ' ==========================================================
    Dim legendRow As Long
    legendRow = r + 2

    wsOut.Cells(legendRow, 1).Value = "Legend"
    wsOut.Cells(legendRow, 1).Font.Bold = True

    wsOut.Cells(legendRow + 1, 1).Value = "C"
    wsOut.Cells(legendRow + 1, 2).Value = "Current / Active"
    wsOut.Cells(legendRow + 1, 1).Interior.Color = RGB(0, 176, 240)

    wsOut.Cells(legendRow + 2, 1).Value = "D"
    wsOut.Cells(legendRow + 2, 2).Value = "Discovery"
    wsOut.Cells(legendRow + 2, 1).Interior.Color = RGB(255, 255, 0)

    wsOut.Cells(legendRow + 3, 1).Value = "P"
    wsOut.Cells(legendRow + 3, 2).Value = "Planning"
    wsOut.Cells(legendRow + 3, 1).Interior.Color = RGB(255, 102, 204)

    wsOut.Range(wsOut.Cells(legendRow, 1), wsOut.Cells(legendRow + 3, 2)).Borders.LineStyle = xlContinuous

    ' ==========================================================
    ' FINAL VIEW SETTINGS
    ' ==========================================================
    wsOut.Activate
    wsOut.Range("G3").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(2, timelineEndCol)).AutoFilter

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = calcMode
    Application.EnableEvents = True

    MsgBox "Aha-style roadmap created in a NEW workbook after refreshing Aha / Power Query data." & vbCrLf & vbCrLf & _
           "Source sheet used: " & wsSrc.Name, vbInformation

    Exit Sub

CleanFail:

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = calcMode
    Application.EnableEvents = True

    MsgBox "The Aha-style roadmap macro stopped with this error:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Your original workbook was not intentionally modified.", vbExclamation

End Sub


Function SheetExists(wb As Workbook, sheetName As String) As Boolean

    Dim ws As Worksheet

    SheetExists = False

    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws

End Function


Function GetStatusColor(statusVal As String) As Long

    Select Case UCase(Trim(statusVal))

        Case "C"
            GetStatusColor = RGB(0, 176, 240)

        Case "D"
            GetStatusColor = RGB(255, 255, 0)

        Case "P"
            GetStatusColor = RGB(255, 102, 204)

        Case Else
            GetStatusColor = RGB(180, 180, 180)

    End Select

End Function


Function GetPrimaryStatus(qDict As Object, startQ As String, quarters As Variant) As String

    Dim q As Variant

    If qDict.exists(startQ) Then
        If Len(Trim(CStr(qDict(startQ)))) > 0 Then
            GetPrimaryStatus = CStr(qDict(startQ))
            Exit Function
        End If
    End If

    For Each q In quarters
        If qDict.exists(CStr(q)) Then
            If Len(Trim(CStr(qDict(CStr(q))))) > 0 Then
                GetPrimaryStatus = CStr(qDict(CStr(q)))
                Exit Function
            End If
        End If
    Next q

    GetPrimaryStatus = ""

End Function


Sub MergeColumnSafe(ws As Worksheet, colNum As Long, startRow As Long, endRow As Long)

    Dim startMerge As Long
    Dim i As Long
    Dim currentVal As String
    Dim startVal As String

    If endRow <= startRow Then Exit Sub

    startMerge = startRow

    For i = startRow + 1 To endRow + 1

        startVal = CStr(ws.Cells(startMerge, colNum).Value)

        If i <= endRow Then
            currentVal = CStr(ws.Cells(i, colNum).Value)
        Else
            currentVal = "__END_OF_DATA__"
        End If

        If currentVal <> startVal Or i > endRow Then

            If i - 1 > startMerge Then

                With ws.Range(ws.Cells(startMerge, colNum), ws.Cells(i - 1, colNum))
                    .Merge
                    .VerticalAlignment = xlCenter
                    .HorizontalAlignment = xlCenter
                    .WrapText = True
                End With

            End If

            startMerge = i

        End If

    Next i

End Sub


Sub QuickSortVariant(arr As Variant, first As Long, last As Long)

    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As String

    i = first
    j = last
    pivot = CStr(arr((first + last) \ 2))

    Do While i <= j

        Do While CStr(arr(i)) < pivot
            i = i + 1
        Loop

        Do While CStr(arr(j)) > pivot
            j = j - 1
        Loop

        If i <= j Then

            temp = CStr(arr(i))
            arr(i) = arr(j)
            arr(j) = temp

            i = i + 1
            j = j - 1

        End If

    Loop

    If first < j Then QuickSortVariant arr, first, j
    If i < last Then QuickSortVariant arr, i, last

End Sub
