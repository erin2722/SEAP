VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub controlData()
    ActiveSheet.Shapes.Range(Array("ImportButton")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Loading..."
    Range("A1").Select
    Application.Wait (Now + TimeValue("0:0:01"))
    importData
End Sub
Public Function importData()
    Dim intChoice As Integer
    Dim strPaths As String
    Dim i As Integer
    
    'gives number of data row being checked
    Dim currentDataRow As Integer
    currentDataRow = 3
    
    'range that while loop will use to check if row is empty
    Dim checkingRow As String
    checkingRow = "A" & CStr(currentDataRow)

    'allows easy switching between wbks
    Dim MainWbk As Workbook
    Dim SourceWbk As Workbook
    Set MainWbk = ActiveWorkbook
    
    'allow the user to select multiple files
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    'determine what choices the user made
    If intChoice <> 0 Then
        Columns("A:W").Clear
        Application.ScreenUpdating = False
        Columns("B:C").Delete
        For i = 1 To Application.FileDialog(msoFileDialogOpen _
            ).SelectedItems.Count
            strPath = Application.FileDialog(msoFileDialogOpen _
            ).SelectedItems(i)
            Set SourceWbk = Workbooks.Open(strPath)
            If i = 1 Then
                SourceWbk.Worksheets("Sheet1").Range("A1:U1500").Copy
            Else
                SourceWbk.Worksheets("Sheet1").Range("A2:U1500").Copy
            End If
            MainWbk.Worksheets("Sheet1").Activate
        
            'figures out the range that data should be pasted to
            While CStr(Range(checkingRow).Value) <> ""
                currentDataRow = currentDataRow + 1
                checkingRow = "A" & CStr(currentDataRow)
            Wend
            'pastes the data
            ActiveSheet.Range("A" & CStr(currentDataRow) & ":U" & CStr(currentDataRow)).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            SourceWbk.Close
        Next i
        sortByTime
        timeElapsed
        Application.ScreenUpdating = True
    End If
    
    ActiveSheet.Shapes.Range(Array("ImportButton")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Import New Data"
    
    'cleans up files for user
    Range("A1").Select
    Application.CutCopyMode = False '
End Function
Function sortByTime()
    'sorts time of day from earliest to latest
    Dim currentDataRow As Integer
    currentDataRow = findCurrentDataRow
    Range("A4:A" & CStr(currentDataRow)).Select
    ActiveWorkbook.Worksheets("sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sheet1").Sort.SortFields.Add Key:=Range("A4"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("sheet1").Sort
        .SetRange Range("A4:W" & CStr(currentDataRow))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Function
Function timeElapsed()
     Dim currentDataRow As Integer
     currentDataRow = findCurrentDataRow
    
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("B:B").Select
    ActiveSheet.Paste
    Selection.NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
    Range("C3").FormulaR1C1 = "Time Elapsed (h:m:s)"
    Range("C4").Select
    Selection.FormulaR1C1 = "=RC[-1]-R4C2"
    Range("C4").NumberFormat = "[h]:mm:ss"
    Selection.AutoFill Destination:=Range("C4:C" & CStr(currentDataRow)), Type:=xlFillDefault
    
    ActiveSheet.Range("B1:B2").Clear

    Range("A1").Select
    Application.CutCopyMode = False
End Function
Sub Graph_O2()
    Dim currentDataRow As Integer
    
    
    
    currentDataRow = findCurrentDataRow
    
     
    'creates chart and adds data
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers, Width:=450).Select
    ActiveChart.SetSourceData Source:=Range("sheet1!$C$4:$C$" & CStr(currentDataRow) & ",sheet1!$D4:$H" & CStr(currentDataRow))
    
    'creates legend
    ActiveChart.FullSeriesCollection(1).Name = "=""Chamber A"""
    ActiveChart.FullSeriesCollection(2).Name = "=""Chamber B"""
    ActiveChart.FullSeriesCollection(3).Name = "=""Chamber C"""
    ActiveChart.FullSeriesCollection(4).Name = "=""Chamber D"""
    ActiveChart.FullSeriesCollection(5).Name = "=""Chamber E"""
    ActiveChart.SetElement (msoElementLegendRight)
    
    'formats graph and adds title with initial timestamp
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Range("A4").Text + " O2 vs Time"
    Selection.Format.TextFrame2.TextRange.Characters.Text = Range("A4").Text + " O2 vs Time"
    'ActiveSheet.ChartObjects(1).Width = 450
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    
    'this is to add labels and format axes
    ActiveChart.ChartArea.Select
    With ActiveChart.Axes(xlPrimary)
     .HasTitle = True
     .AxisTitle.Text = "Elapsed Time (H:M:S)"
     .MaximumScale = CDbl(Range("C" & currentDataRow))
    End With
    With ActiveChart.Axes(xlSecondary)
     .HasTitle = True
     .AxisTitle.Text = "%O2"
    End With
End Sub
Sub Graph_CO2()
    Dim currentDataRow As Integer
    
    currentDataRow = findCurrentDataRow
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers, Width:=450).Select
    ActiveChart.SetSourceData Source:=Range("sheet1!$C$4:$C$" & CStr(currentDataRow) & ",sheet1!$I4:$M" & CStr(currentDataRow))
    
    ActiveChart.FullSeriesCollection(1).Name = "=""Chamber A"""
    ActiveChart.FullSeriesCollection(2).Name = "=""Chamber B"""
    ActiveChart.FullSeriesCollection(3).Name = "=""Chamber C"""
    ActiveChart.FullSeriesCollection(4).Name = "=""Chamber D"""
    ActiveChart.FullSeriesCollection(5).Name = "=""Chamber E"""
    ActiveChart.SetElement (msoElementLegendRight)
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Range("A4").Text + " CO2 vs Time"
    Selection.Format.TextFrame2.TextRange.Characters.Text = Range("A4").Text + " CO2 vs Time"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With

    ActiveChart.ChartArea.Select
    With ActiveChart.Axes(xlPrimary)
     .HasTitle = True
     .AxisTitle.Text = "Elapsed Time (H:M:S)"
     .MaximumScale = CDbl(Range("C" & currentDataRow))
    End With
    With ActiveChart.Axes(xlSecondary)
     .HasTitle = True
     .AxisTitle.Text = "CO2 (ppm)"
    End With
End Sub
Sub Graph_Temp()
    Dim currentDataRow As Integer
    
    currentDataRow = findCurrentDataRow
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers, Width:=450).Select
    ActiveChart.SetSourceData Source:=Range("sheet1!$C$4:$C$" & CStr(currentDataRow) & ",sheet1!$N4:$R" & CStr(currentDataRow))
    
    ActiveChart.FullSeriesCollection(1).Name = "=""Chamber A"""
    ActiveChart.FullSeriesCollection(2).Name = "=""Chamber B"""
    ActiveChart.FullSeriesCollection(3).Name = "=""Chamber C"""
    ActiveChart.FullSeriesCollection(4).Name = "=""Chamber D"""
    ActiveChart.FullSeriesCollection(5).Name = "=""Chamber E"""
    ActiveChart.SetElement (msoElementLegendRight)
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Range("A4").Text + " Temp vs Time"
    Selection.Format.TextFrame2.TextRange.Characters.Text = Range("A4").Text + " Temp vs Time"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    'this is to add labels
    ActiveChart.ChartArea.Select
    With ActiveChart.Axes(xlPrimary)
     .HasTitle = True
     .AxisTitle.Text = "Elapsed Time (H:M:S)"
     .MaximumScale = CDbl(Range("C" & currentDataRow))
    End With
    With ActiveChart.Axes(xlSecondary)
     .HasTitle = True
     .AxisTitle.Text = "Temperature (F)"
    End With
End Sub
Sub Graph_Depth()
    Dim currentDataRow As Integer
    
    currentDataRow = findCurrentDataRow
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers, Width:=450).Select
    ActiveChart.SetSourceData Source:=Range("sheet1!$C$4:$C$" & CStr(currentDataRow) & ",sheet1!$S4:$W" & CStr(currentDataRow))
    ActiveChart.FullSeriesCollection(1).Name = "=""Chamber A"""
    ActiveChart.FullSeriesCollection(2).Name = "=""Chamber B"""
    ActiveChart.FullSeriesCollection(3).Name = "=""Chamber C"""
    ActiveChart.FullSeriesCollection(4).Name = "=""Chamber D"""
    ActiveChart.FullSeriesCollection(5).Name = "=""Chamber E"""
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = Range("A4").Text + " Depth vs Time"
    Selection.Format.TextFrame2.TextRange.Characters.Text = Range("A4").Text + " Depth vs Time"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    'this is to add labels
    ActiveChart.ChartArea.Select
    With ActiveChart.Axes(xlPrimary)
     .HasTitle = True
     .AxisTitle.Text = "Elapsed Time (H:M:S)"
     .MaximumScale = CDbl(Range("C" & currentDataRow))
    End With
    With ActiveChart.Axes(xlSecondary)
     .HasTitle = True
     .AxisTitle.Text = "Depth (fsw)"
    End With
End Sub
Sub findDanger()
    Dim currentDataRow As Integer
    
    Dim O2LowWarning, O2HighWarning, CO2LowWarning, CO2HighWarning, O2LowAlarm, O2HighAlarm, CO2LowAlarm, CO2HighAlarm, depthHigh As Integer
    
    
    O2LowWarning = 0.44
    O2HighWarning = 0.48
    O2LowAlarm = 0.42
    O2HighAlarm = 0.6
    
    CO2HighWarning = 0.5
    CO2HighAlarm = 2
    
    depthHighWarning = 1000
    
    Range("Z15").FormulaR1C1 = CDbl(Range("Z7").Value)
    
    currentDataRow = findCurrentDataRow
    ' put values in this loop to find O2 danger values
    For Each cell In Range("D4:H" & CStr(currentDataRow))
        If CDbl(cell.Value) > O2HighWarning Then
            If CDbl(cell.Value) >= O2HighAlarm Then
                'alarm
                cell.Interior.Color = 128
            Else
                'warning
                cell.Interior.Color = 52479
            End If
        End If
        If CDbl(cell.Value) < O2LowWarning Then
            If CDbl(cell.Value) <= O2LowAlarm Then
                'alarm
                cell.Interior.Color = 255
            Else
                'warning
                cell.Interior.Color = 10092543
            End If
        End If
    Next
    'this is for CO2
    For Each cell In Range("I4:M" & CStr(currentDataRow))
        If CDbl(cell.Value) > CO2HighWarning Then
            If CDbl(cell.Value) >= CO2HighAlarm Then
                'alarm
                cell.Interior.Color = 128
            Else
                'warning
                cell.Interior.Color = 52479
            End If
        End If
    Next
    'this is for depth
    For Each cell In Range("S4:W" & CStr(currentDataRow))
        If CDbl(cell.Value) >= depthHighWarning Then
            'warning
            cell.Interior.Color = 52479
        End If
    Next
End Sub
Public Function findCurrentDataRow() As Integer
    'gives number of data row being checked
     Dim currentDataRow As Integer
     currentDataRow = 3
     'range that while loop will use to check if row is empty
     Dim checkingRow As String
     checkingRow = "A" & CStr(currentDataRow)
     While CStr(Range(checkingRow).Value) <> ""
        currentDataRow = currentDataRow + 1
        checkingRow = "A" & CStr(currentDataRow)
    Wend
    findCurrentDataRow = currentDataRow - 1
    Exit Function
End Function
