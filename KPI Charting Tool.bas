Attribute VB_Name = "ChartingTool"
Public title As String
Public improvement1 As String
Public improvement2 As String
Public improvement3 As String
Public improvement4 As String
Public improvement5 As String
Public mm1 As Integer
Public mm2 As Integer
Public mm3 As Integer
Public mm4 As Integer
Public mm5 As Integer
Public yyyy1 As Integer
Public yyyy2 As Integer
Public yyyy3 As Integer
Public yyyy4 As Integer
Public yyyy5 As Integer

Sub Show_User_Form()
Attribute Show_User_Form.VB_ProcData.VB_Invoke_Func = "d\n14"

KPIChartingTool.Show

End Sub

Sub create_report_now()

Application.DisplayAlerts = False

Dim Table As ListObject
Dim Rng As Range

'On Error GoTo NewWorksheetError
With ActiveSheet
        ShName = "formatteddata"
    .Copy After:=Sheets(Worksheets.Count)
End With
Sheets(Worksheets.Count).Name = ShName
'On Error GoTo 0

Set Rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
'On Error GoTo GraphOverlap
    Set Table = ActiveSheet.ListObjects.Add(xlSrcRange, Rng, , xlYes)
'On Error GoTo 0
Table.TableStyle = "TableStyleMedium15"

With Table.Sort
    .SortFields.clear
    .SortFields.Add Key:=Range("A1"), Order:=xlAscending
    .Apply
End With

With Table
    .ListColumns(8).Delete
    .ListColumns.Add(8).Name = "Vertical Line"
    .ListColumns.Add(9).Name = "Letter"
    .ListColumns.Add(10).Name = "Formatted Dates"
  ' .ListColumns(2).Range.Copy Destination:=ActiveSheet.Range("J1")
   '.ListColumns(2).DataBodyRange.NumberFormat = "mmm"
    '.ListColumns(10).DataBodyRange.NumberFormat = "mm-yyyy"
End With


Dim i As Integer
Dim cell As Range
Dim Val As String
For i = 2 To Table.DataBodyRange.Rows.Count
    Val = Format(Cells(i, 2), "mmm-yy")
    Cells(i, 10).NumberFormat = "@"
    Cells(i, 10).Value = Val
Next i

Dim months(1 To 5) As Integer
months(1) = mm1
months(2) = mm2
months(3) = mm3
months(4) = mm4
months(5) = mm5

Dim years(1 To 5) As Integer
years(1) = yyyy1
years(2) = yyyy2
years(3) = yyyy3
years(4) = yyyy4
years(5) = yyyy5

Dim letters(1 To 5) As Variant
letters(1) = "A"
letters(2) = "B"
letters(3) = "C"
letters(4) = "D"
letters(5) = "E"

FirstRow = 2
'On Error GoTo SIRFormat
Range1End = Range("A:A").Find("1", searchdirection:=xlPrevious, LookAt:=xlWhole).Row
Range3End = Range("A:A").Find("3", searchdirection:=xlPrevious, LookAt:=xlWhole).Row
Range6End = Range("A:A").Find("6", searchdirection:=xlPrevious, LookAt:=xlWhole).Row
Range12End = Range("A:A").Find("12", searchdirection:=xlPrevious, LookAt:=xlWhole).Row
Range3Start = Range1End + 1
Range6Start = Range3End + 1
Range12Start = Range6End + 1

Range1SIR = Range(Cells(FirstRow, 4), Cells(Range1End, 4))
Range3SIR = Range(Cells(Range3Start, 4), Cells(Range3End, 4))
Range6SIR = Range(Cells(Range6Start, 4), Cells(Range6End, 4))
Range12SIR = Range(Cells(Range12Start, 4), Cells(Range12End, 4))
'On Error GoTo 0

Dim Row As Integer
Dim currentMonth As Variant
Dim currentYear As Variant
Count = 1

For Row = 1 To Range1End Step 1
        currentMonth = month(Table.DataBodyRange.Cells(Row, 2).Value)
        currentYear = year(Table.DataBodyRange.Cells(Row, 2).Value)
        If ((currentMonth = months(Count)) And (currentYear = years(Count))) Then
            Table.DataBodyRange.Cells(Row, 8).Value = 100
            Table.DataBodyRange.Cells(Row, 9).Value = letters(Count)
            Count = Count + 1
            If (Count > 5) Then
                Exit For
            End If
        End If
Next Row


xAxis1 = Range(Cells(FirstRow, 10), Cells(Range1End, 10))
xAxis2 = Range(Cells(FirstRow, 9), Cells(Range1End, 9))
labelData = Range(Cells(FirstRow, 8), Cells(Range1End, 8))
Labels = Range(Cells(FirstRow, 9), Cells(Range1End, 9))

Set NewWs = ActiveWorkbook.Worksheets.Add(Type:=xlWorksheet)
NewWs.Name = title

Dim MyChart As ChartObject
Set MyChart = Worksheets(title).ChartObjects.Add(Left:=25, Width:=700, Top:=10, Height:=400)
MyChart.Activate
ActiveChart.Axes(xlCategory).CategoryType = xlCategoryScale
Set NewS = ActiveChart.SeriesCollection.NewSeries
With NewS
    .Values = Range1SIR
    .Name = "1 Month"
    .XValues = xAxis1
    .AxisGroup = xlPrimary
    .ChartType = xlLineMarkers
    .Format.Fill.ForeColor.RGB = RGB(70, 150, 255)
    .Format.Line.ForeColor.RGB = RGB(70, 150, 255)
End With

Set NewS = ActiveChart.SeriesCollection.NewSeries
With NewS
    .Values = Range3SIR
    .Name = "3 Month"
    .XValues = xAxis1
    .AxisGroup = xlPrimary
    .ChartType = xlLineMarkers
    .Format.Fill.ForeColor.RGB = RGB(255, 119, 0)
    .Format.Line.ForeColor.RGB = RGB(255, 119, 0)
End With

Set NewS = ActiveChart.SeriesCollection.NewSeries
With NewS
    .Values = Range6SIR
    .Name = "6 Month"
    .XValues = xAxis1
    .AxisGroup = xlPrimary
    .ChartType = xlLineMarkers
    .Format.Fill.ForeColor.RGB = RGB(150, 150, 150)
    .Format.Line.ForeColor.RGB = RGB(150, 150, 150)
End With

Set NewS = ActiveChart.SeriesCollection.NewSeries
With NewS
    .Values = Range12SIR
    .ChartType = xlLineMarkers
    .Name = "12 Month"
    .XValues = xAxis1
    .AxisGroup = xlPrimary
    .Format.Fill.ForeColor.RGB = RGB(255, 200, 0)
    .Format.Line.ForeColor.RGB = RGB(255, 200, 0)
End With

Set NewS = ActiveChart.SeriesCollection.NewSeries
With NewS
    .Values = labelData
    .ChartType = xlColumnStacked
    .Name = "labels"
    .XValues = Labels
    .AxisGroup = xlSecondary
    .Format.Fill.ForeColor.RGB = RGB(0, 0, 0)

End With


With ActiveChart
    .HasTitle = True
    .ChartTitle.Characters.Text = title
    
    'Y Axis
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Adjusted SIR"
    .Axes(xlValue).MaximumScale = 50
    
    'X Axis
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Manufacturing Month"
    .Axes(xlCategory, xlPrimary).TickLabels.Orientation = 90
    .Axes(xlCategory, xlPrimary).Select
    Selection.TickLabels.Orientation = xlUpward
   
    'Secondary Axes
    .HasAxis(xlCategory, xlSecondary) = True
    
    'Secondary Y Axis
    .HasAxis(xlValue, xlSecondary) = False
    
    'Secondary X Axis
    .Axes(xlCategory, xlSecondary).MajorTickMark = xlTickMarkNone
    .Axes(xlCategory, xlSecondary).TickLabels.Font.Size = 18
    
    .Legend.LegendEntries(1).Delete
    .SetElement (msoElementLegendBottom)
    .ChartGroups(1).GapWidth = 500
    .FullSeriesCollection(5).Select
End With

'Pattern for lines

'With Selection.Format.Fill
      '  .Visible = msoTrue
       ' .Patterned msoPattern50Percent
'End With

ActiveChart.PlotArea.Select

With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.0500000007
        .Transparency = 0
        .Solid
End With

Dim projectsTable As ListObject
Set projectsTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range("R5:S10"))
projectsTable.TableStyle = "TableStyleMedium2"

projectsTable.DataBodyRange(1, 1).Value = "A"
projectsTable.DataBodyRange(1, 2).Value = improvement1

projectsTable.DataBodyRange(2, 1).Value = "B"
projectsTable.DataBodyRange(2, 2).Value = improvement2

projectsTable.DataBodyRange(3, 1).Value = "C"
projectsTable.DataBodyRange(3, 2).Value = improvement3

projectsTable.DataBodyRange(4, 1).Value = "D"
projectsTable.DataBodyRange(4, 2).Value = improvement4

projectsTable.DataBodyRange(5, 1).Value = "E"
projectsTable.DataBodyRange(5, 2).Value = improvement5

With Worksheets(title).Columns("R")
    .ColumnWidth = .ColumnWidth * 0.5
End With

With Worksheets(title).Columns("S")
    .ColumnWidth = .ColumnWidth * 3
End With

Sheets("formatteddata").Delete


'NewWorksheetError:
  '  ShName = "formatteddata2"
'SIRFormat:
 '   MsgBox ("Error: this data set does not have correct SIR formatting. Please use a dataset that contains 1,3,6,12 month SIR")
  '  End
'GraphOverlap:
  '  MsgBox ("Error: the workseet you have selected is not formatted correctly. Are you running this macro on the right worksheet???")
  '  End

End Sub


