Attribute VB_Name = "Module1"
'CHART-CREATING MACRO DESIGNED FOR SURVEYMONKEY OUTPUTS
'(C) 2014 FRIDA HATAMI - ANNENBERG FOUNDATION
'FRIDAHATAMI@GMAIL.COM, 818-613-9651
'-------------------------------CHANGELOG-------------------------------
'v1.1 - FEBRUARY 9, 2014
'NEW NORMALIZATION METHOD FOR DATA
'NEW COLUMNS INSERTED FOR NORMALIZED DATA
'UPDATED CHARTING
'----------------------------------------------------------------------
'v1.0 - FEBRUARY 9, 2014
'INSERTS 2 NEW COLUMNS BASED ON MAX OF 5 ANSWER OPTION QUESTIONS
'CHANGES COLUMN WIDTHS FOR CATEGORIES, ANSWER OPTIONS, AND RATING/RESPONSE
'INSERTS FORMULAS FOR SUMMATION OF TOP TWO ANSWER CHOICES AND % OF TOTAL
'UNMERGES (AND LATER REMERGES) CATEGORIES IN COLUMN 1 TO AID IN AXIS LABELING
'CREATES CHARTS FOR QUESTIONS THAT HAVE 3 OR MORE ANSWER OPTIONS
'   - TITLE IS SET AS QUESTION # ONLY
'   - DATA IS SORTED FIRST BY % OF TOTAL, THEN BY TOP ANSWER CHOICE
'   - SERIES COLORS ARE SET IN THE FUNCTION (v1.0 - ORANGE TINT)
'   - CHART WIDTH IS SET TO 400, WHILE HEIGHT IS DYNAMICALLY ASSIGNED
'     BASED ON NUMBER OF CATEGORIES FOR THE QUESTION
'   - MAJOR AXIS LENGHT IS DYNAMICALLY SET BASED ON THE MAXIMUM OF TOP TWO ANSWERS
'   - SERIES DATA LABELS ARE ON, AND LEGEND ENTRIES ARE BASED ON TABLE HEADERS (TOP 2 ANSWERS)
'   - CHART BORDER IS SET TO NONE
'----------------------------------------------------------------------

Sub main()
    Dim sh As Worksheet, val As Integer, qNo As Integer
    
    Set sh = ActiveSheet
    
    sh.Cells.RowHeight = 20
    
    val = Application.WorksheetFunction.CountIf(Range("A1:A1000"), "Answer Options") 'FIND TOTAL # OF QUESTIONS
    
    qArr = allQs(val, sh)
    
    Columns("H:K").Insert
    ActiveSheet.Columns("A:B").ColumnWidth = 15
    ActiveSheet.Columns("C:K").ColumnWidth = 5
    ActiveSheet.Columns("L:M").ColumnWidth = 12
    
    Call sumAndPerc(val, qArr)
    
    'UNMERGE CELLS IN FIRST COLUMN FOR AXIS LABELING
    For i = 1 To val
        For j = qArr(i, 1) + 1 To qArr(i, 2) - 1
            Cells(j, 1).MergeCells = False
        Next j
    Next i
    
    'CREATE CHARTS
    For i = 1 To val
        qNo = i
        Call createChart(qArr, qNo, sh)
    Next i
    
    'MERGE CELLS IN FIRST COLUMN AFTER CHARTS ARE DONE
    Application.DisplayAlerts = False
    For i = 1 To val
        For j = qArr(i, 1) + 1 To qArr(i, 2) - 1
            Range(Cells(j, 1), Cells(j, 2)).MergeCells = True
        Next j
    Next i
    Application.DisplayAlerts = True
    
    
End Sub

'----------------------------------------------------------------------
'FUNCTION FOR CREATING CHARTS FOR QUESTIONS WITH MORE THAN 2 OPTIONS
'IN IT'S CURRENT FORM, THE FUNCTION IS DESIGNED TO SORT 5 AND 3 CATEGORY
'QUESTIONS ONLY.  IF OTHER TYPES OF QUESTIONS ARE PRESENT, FUNCTION WILL
'NEED TO BE MODIFIED
'----------------------------------------------------------------------
Public Function createChart(qArr As Variant, qNo As Integer, sh As Worksheet)
    s = qArr(qNo, 1) + 1
    e = qArr(qNo, 2) - 1
    If qArr(qNo, 3) = 5 Then 'FIVE CATEGORY QUESTION SORTING
        With Range(Cells(s - 1, 1), Cells(e, 13))
            .Sort Key1:=Range(Cells(s, 11), Cells(e, 11)), _
            Order1:=xlDescending, _
            Key2:=Range(Cells(s, 10), Cells(e, 10)), _
            Order2:=xlDescending, _
            Header:=xlGuess
        End With
        Set catRange = Range(Cells(s, 1), Cells(e, 1))
        Set valRange = Application.Union(Range(Cells(s, 8), Cells(e, 9)), Range(Cells(s, 11), Cells(e, 11)))
    ElseIf qArr(qNo, 3) = 3 Then 'THREE CATEGORY QUESTION SORTING
        With Range(Cells(s - 1, 1), Cells(e, 11))
            .Sort Key1:=Range(Cells(s, 9), Cells(e, 9)), _
            Order1:=xlDescending, _
            Key2:=Range(Cells(s, 8), Cells(e, 8)), _
            Order2:=xlDescending, _
            Header:=xlGuess
        End With
        Set catRange = Range(Cells(s, 1), Cells(e, 1))
        Set valRange = Application.Union(Range(Cells(s, 6), Cells(e, 7)), Range(Cells(s, 9), Cells(e, 9)))
    End If
    If qArr(qNo, 3) > 2 Then
        sh.Shapes.AddChart(xlBarStacked).Select
        With ActiveChart
            .SetSourceData Source:=Union(catRange, valRange)
            .Axes(xlCategory).ReversePlotOrder = True
            .Axes(xlCategory).TickLabels.Font.Size = 10
            .HasAxis(xlSecondary) = False
            .Axes(xlSecondary).HasMajorGridlines = False
            '.Axes(xlSecondary).MaximumScale = WorksheetFunction.Max(Range(Cells(s, qArr(qNo, 3) + 3), Cells(e, qArr(qNo, 3) + 3))) * 1.25
            .Axes(xlSecondary).MaximumScale = 120
            .ChartArea.Border.LineStyle = xlNone
            .Parent.Top = Cells(qArr(qNo, 1) - 1, 15).Top
            .Parent.Left = Cells(qArr(qNo, 1) - 1, 15).Left
            .ChartArea.Width = 400
            .ChartArea.Height = Range(Cells(qArr(qNo, 1) - 1, 1), Cells(qArr(qNo, 2) + 2, 1)).Height
            .HasTitle = True
            .ChartTitle.Text = "Question " & Left(Cells(qArr(qNo, 1) - 1, 1).Value, InStr(1, Cells(qArr(qNo, 1) - 1, 1).Value, ".") - 1)
            .ChartTitle.Font.FontStyle = "Bold"
            .Legend.Font.Size = 11
            .Legend.Position = xlLegendPositionBottom
            For j = 1 To 3
                If j < 3 Then
                    '.FullSeriesCollection(j).ApplyDataLabels
                    '.FullSeriesCollection(j).DataLabels.Font.Size = 10
                    .SeriesCollection(j).Name = Cells(qArr(qNo, 1), qArr(qNo, 3) + j).Value
                Else
                    .SeriesCollection(j).ApplyDataLabels
                    .SeriesCollection(j).DataLabels.Font.Size = 12
                    .SeriesCollection(j).DataLabels.Font.FontStyle = "Bold"
                    .SeriesCollection(j).DataLabels.Position = xlLabelPositionInsideBase
                    .SeriesCollection(j).Format.Fill.Visible = msoFalse
                    .Legend.LegendEntries(j).Delete
                End If
            Next j
            ActiveChart.SeriesCollection(1).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 218, 193) '<---CHANGE COLOR FOR FIRST SERIES HERE
                .Transparency = 0
                .Solid
            End With
            ActiveChart.SeriesCollection(2).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 157, 91) '<---CHANGE COLOR FOR SECOND SERIES HERE
                .Transparency = 0
                .Solid
            End With
        End With
    End If
End Function

'----------------------------------------------------------------------
'FUNCTION FOR FINDING THE BEGINNING AND ENDING OF ALL QUESTIONS
'THIS FUNCTION CREATES AN (N BY 3) ARRAY WITH THE FOLLOWING FORMAT
'-------------------------------------
'| Q start | Q end | # of categories |
'|    .    |   .   |        .        |
'|    .    |   .   |        .        |
'-------------------------------------
'----------------------------------------------------------------------
Public Function allQs(qs As Integer, sh As Worksheet) As Variant
    i = 1
    str1 = "Answer Options"
    str2 = "answered question"
    str3 = "Response Count"
    Dim x()
    ReDim x(1 To qs, 1 To 3)
    For Each Cell In Range("A1:A1000")
        If Cell.Value = str1 Then
            x(i, 1) = Cell.Row
            Cell.Select
            If ActiveCell.Offset(0, 7) = str3 Then
                x(i, 3) = 5
            ElseIf ActiveCell.Offset(0, 5) = str3 Then
                x(i, 3) = 3
            ElseIf ActiveCell.Offset(0, 1) = str3 Then
                x(i, 3) = 0
            End If
        ElseIf Cell.Value = str2 Then
            x(i, 2) = Cell.Row
            i = i + 1
        End If
    Next
    allQs = x
    'UNCOMMENT LINE BELOW TO SEE ARRAY SAMPLE IN CELLS Q1:S3
    'Range("Q1:S3") = x
End Function

'----------------------------------------------------------------------
'FUNCTION FOR CHANGING ROW HEADERS OF THE INSERTED COLUMNS
'AND APPLYING FORMULAS FOR CELLS IN THE NEW COLUMNS FOR SUMMATION
'AND PERCENTAGE CALCULATION OF TOP TWO RESPONSES
'v1.1 ADDS TWO MORE COLUMNS WITH NORMALIZED DATA FOR PLOTTING
'----------------------------------------------------------------------
Public Function sumAndPerc(qs As Integer, qArr As Variant)
    ampStr = "," & Chr(34) & " & " & Chr(34) & ","
    normStr = Chr(34) & "Norm. " & Chr(34) & ","
    For i = 1 To qs
        If qArr(i, 3) = 5 Then '5 CATEGORY QUESTIONS
            Cells(qArr(i, 1), 8).Value = "=CONCATENATE(" & normStr & Cells(qArr(i, 1), 6).Address & ")"
            Cells(qArr(i, 1), 9).Value = "=CONCATENATE(" & normStr & Cells(qArr(i, 1), 7).Address & ")"
            Cells(qArr(i, 1), 10).Value = "=CONCATENATE(" & Cells(qArr(i, 1), 6).Address & ampStr & Cells(qArr(i, 1), 7).Address & ")"
            Cells(qArr(i, 1), 11).Value = "% of Total"
            For j = 1 To qArr(i, 2) - qArr(i, 1) - 1
                cr = qArr(i, 1) + j
                Cells(cr, 8).Value = "=ROUND((F" & cr & "/M" & cr & ")*100,0)"
                Cells(cr, 9).Value = "=ROUND((G" & cr & "/M" & cr & ")*100,0)"
                Cells(cr, 10).Value = "=SUM(F" & cr & ":G" & cr & ")"
                Cells(cr, 11).Value = "=ROUND(J" & cr & "/M" & cr & ",2)"
                Cells(cr, 11).NumberFormat = "0%"
            Next j
        ElseIf qArr(i, 3) = 3 Then '3 CATEGORY QUESTIONS
            'FIRST WE MOVE THE TWO COLUMNS OVER
            Range(Cells(qArr(i, 1), 6), Cells(qArr(i, 2) - 1, 7)).Select
            Selection.Copy
            Cells(qArr(i, 1), 10).Select
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            'THEN INSERT FORMULAS FOR NEW COLUMNS
            Cells(qArr(i, 1), 6).Value = "=CONCATENATE(" & normStr & Cells(qArr(i, 1), 4).Address & ")"
            Cells(qArr(i, 1), 7).Value = "=CONCATENATE(" & normStr & Cells(qArr(i, 1), 5).Address & ")"
            Cells(qArr(i, 1), 8).Value = "=CONCATENATE(" & Cells(qArr(i, 1), 4).Address & ampStr & Cells(qArr(i, 1), 5).Address & ")"
            Cells(qArr(i, 1), 9).Value = "% of Total"
            For j = 1 To qArr(i, 2) - qArr(i, 1) - 1
                cr = qArr(i, 1) + j
                Cells(cr, 6).Value = "=ROUND((D" & cr & "/K" & cr & ")*100,0)"
                Cells(cr, 7).Value = "=ROUND((E" & cr & "/K" & cr & ")*100,0)"
                Cells(cr, 8).Value = "=SUM(D" & cr & ":E" & cr & ")"
                Cells(cr, 9).Value = "=(H" & cr & "/K" & cr & ")"
                Cells(cr, 9).NumberFormat = "0%"
            Next j
        End If
    Next i
End Function



