Option Explicit

Sub BuildOperationsRoadmap()

    ' ==============================
    ' DECLARE VARIABLES
    ' ==============================
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim colMap As Object, dict As Object, qDict As Object
    
    Dim lastCol As Long, lastRow As Long
    Dim data As Variant
    
    Dim colCategory As Long, colSupporting As Long, colAlignment As Long
    Dim colName As Long, colQuarter As Long, colStatus As Long
    
    Dim quarters As Variant
    Dim keys As Variant
    
    Dim i As Long, c As Long, r As Long, colIndex As Long
    
    Dim cat As String, supp As String, align As String, nameVal As String
    Dim qtrName As String, statusVal As String, key As String
    
    Dim prevCat As String, prevSupp As String, prevAlign As String
    Dim currCat As String, currSupp As String, currAlign As String, currName As String
    
    Dim k As Variant, qCol As Variant
    Dim parts() As String
    
    ' ==============================
    ' PERFORMANCE SETTINGS
    ' ==============================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsSrc = ThisWorkbook.Sheets("Query1")

    ' ==============================
    ' RESET OUTPUT
    ' ==============================
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Operations_Formatted_Output").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = Worksheets.Add
    wsOut.Name = "Operations_Formatted_Output"

    ' ==============================
    ' QUARTERS
    ' ==============================
    quarters = Array( _
        "Q.FY26 Q1", "Q.FY26 Q2", "Q.FY26 Q3", "Q.FY26 Q4", _
        "Q.FY27 Q1", "Q.FY27 Q2", "Q.FY27 Q3", "Q.FY27 Q4", _
        "Q.FY28 Q1", "Q.FY28 Q2", "Q.FY28 Q3", "Q.FY28 Q4")

    ' ==============================
    ' MAP COLUMNS
    ' ==============================
    Set colMap = CreateObject("Scripting.Dictionary")
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        colMap(Trim(wsSrc.Cells(1, c).Value)) = c
    Next c

    colCategory = colMap("MFS Category")
    colSupporting = colMap("Supporting Objectives")
    colAlignment = colMap("Initiative Alignment")
    colName = colMap("Initiative name")
    colQuarter = colMap("Quarter")
    colStatus = colMap("Status")

    ' ==============================
    ' LOAD DATA
    ' ==============================
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    data = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Value

    ' ==============================
    ' BUILD MODEL
    ' ==============================
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(data, 1)

        cat = Trim(data(i, colCategory))
        supp = Trim(data(i, colSupporting))
        align = Trim(data(i, colAlignment))
        nameVal = Trim(data(i, colName))
        qtrName = Trim(data(i, colQuarter))
        statusVal = UCase(Trim(data(i, colStatus)))

        If statusVal <> "C" And statusVal <> "D" And statusVal <> "P" Then GoTo SkipRow

        key = cat & "|" & supp & "|" & align & "|" & nameVal

        If Not dict.exists(key) Then
            
            Set qDict = CreateObject("Scripting.Dictionary")
            
            For Each qCol In quarters
                qDict(qCol) = ""
            Next qCol
            
            dict.Add key, Array(cat, supp, align, nameVal, qDict)
        End If

        If dict(key)(4).exists(qtrName) Then
            dict(key)(4)(qtrName) = statusVal
        End If

SkipRow:
    Next i

    keys = dict.keys
    If dict.Count > 1 Then QuickSort keys, LBound(keys), UBound(keys)

    ' ==============================
    ' HEADERS
    ' ==============================
    wsOut.Cells(1, 1) = "MFS Category"
    wsOut.Cells(1, 2) = "Supporting Initiative"
    wsOut.Cells(1, 3) = "Initiative Alignment"
    wsOut.Cells(1, 4) = "Initiative Name"

    colIndex = 5
    For Each qCol In quarters
        wsOut.Cells(1, colIndex) = qCol
        colIndex = colIndex + 1
    Next qCol

    wsOut.Rows(1).Font.Bold = True

    ' ==============================
    ' WRITE DATA
    ' ==============================
    r = 2

    For Each k In keys

        parts = Split(k, "|")

        currCat = parts(0)
        currSupp = parts(1)
        currAlign = parts(2)
        currName = parts(3)

        wsOut.Cells(r, 1) = currCat
        wsOut.Cells(r, 2) = currSupp
        wsOut.Cells(r, 3) = currAlign
        wsOut.Cells(r, 4) = currName

        Set qDict = dict(k)(4)

        colIndex = 5
        For Each qCol In quarters
            wsOut.Cells(r, colIndex) = qDict(qCol)
            colIndex = colIndex + 1
        Next qCol

        r = r + 1

    Next k

    ' ==============================
    ' MERGE CELLS (CATEGORY / SUPPORTING / ALIGNMENT)
    ' ==============================
    Call MergeColumn(wsOut, 1, 2, r - 1)
    Call MergeColumn(wsOut, 2, 2, r - 1)
    Call MergeColumn(wsOut, 3, 2, r - 1)

    ' ==============================
    ' COLOR RULES
    ' ==============================
    Dim dataRange As Range
    Set dataRange = wsOut.Range(wsOut.Cells(2, 5), wsOut.Cells(r - 1, 16))

    With dataRange
        
        ' C = Blue (Currently working)
        .FormatConditions.Add xlCellValue, xlEqual, """C"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 176, 240)
        
        ' D = Yellow (Discovery)
        .FormatConditions.Add xlCellValue, xlEqual, """D"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)
        
        ' P = Pink (Planning)
        .FormatConditions.Add xlCellValue, xlEqual, """P"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 102, 204)

    End With

    ' ==============================
    ' FINAL FORMAT
    ' ==============================
    wsOut.Columns.AutoFit
    wsOut.Range("E2").Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox " Roadmap rebuilt with merges + colors!", vbInformation

End Sub


' ==============================
' MERGE FUNCTION
' ==============================
Sub MergeColumn(ws As Worksheet, colNum As Long, startRow As Long, endRow As Long)

    Dim startMerge As Long, i As Long
    
    startMerge = startRow
    
    For i = startRow + 1 To endRow + 1
        
        If ws.Cells(i, colNum).Value <> ws.Cells(startMerge, colNum).Value Or i > endRow Then
            
            If i - 1 > startMerge Then
                ws.Range(ws.Cells(startMerge, colNum), ws.Cells(i - 1, colNum)).Merge
                ws.Range(ws.Cells(startMerge, colNum), ws.Cells(i - 1, colNum)).VerticalAlignment = xlCenter
            End If
            
            startMerge = i
            
        End If
        
    Next i

End Sub


Sub QuickSort(arr As Variant, first As Long, last As Long)

    Dim i As Long, j As Long
    Dim pivot As String, temp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSort arr, first, j
    If i < last Then QuickSort arr, i, last

End Sub
``
