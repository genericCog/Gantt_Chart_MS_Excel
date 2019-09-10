Option Explicit
   
Private Sub Worksheet_Calculate()
Call DeleteTextboxes
Call ProjectGanttDisplay
End Sub
Private Sub DeleteTextboxes()
    Dim i As Integer
    For i = Shapes.Count To 1 Step -1
         Shapes(i).Delete
    Next i
End Sub

Private Sub ProjectGanttDisplay()
    'look at each cell in the column and compare to each cell in the row
    'if the month and year are the same, display a rectangle over the matching columns
    Call DeleteTextboxes
    Call PageInstructions
    Dim dbl_Counter As Double
    Dim lng_MonthX, lng_YearX, lng_MonthY, lng_YearY, lng_ProjectDuration As Long
    Dim shp_ProjectInfoDisplay, shp_MilestoneMarker, shp_ProjectRectangle As Excel.Shape
    Dim vrt_HorzDate, vrt_VertDate, vrt_ProjectName, vrt_ProjectStart, vrt_ProjectEnd As Variant
    Dim rng_CompareRow, rng_CompareCol As Range
    Dim x As Variant, y As Variant
    Dim str_MilestoneInfo As String
    str_MilestoneInfo = "Milestone 1 & 2" & Chr(10) & "Start" & Chr(10) & "Design Approval"
    Set rng_CompareRow = Range("G3:AN3")
    Set rng_CompareCol = Range("C5:C17")
    For Each x In rng_CompareCol
        lng_MonthX = Month(x)
        lng_YearX = Year(x)
        For Each y In rng_CompareRow
            lng_MonthY = Month(y)
            lng_YearY = Year(y)
            If lng_MonthX = lng_MonthY And lng_YearX = lng_YearY Then
                vrt_HorzDate = y.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                vrt_VertDate = x.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                vrt_ProjectName = CStr(x.Offset(0, -1).Value)
                vrt_ProjectStart = x.Value
                vrt_ProjectEnd = x.Offset(0, 1).Value
                lng_ProjectDuration = (x.Offset(0, 2).Value)
                dbl_Counter = 2 + dbl_Counter
                Set shp_ProjectRectangle = Shapes.AddShape(msoShapeRectangle, _
                Range(vrt_HorzDate).Offset(dbl_Counter, 0).Left, _
                Range(vrt_HorzDate).Offset(dbl_Counter, 0).Top, _
                lng_ProjectDuration * Range(vrt_HorzDate).Offset(dbl_Counter, lng_ProjectDuration).EntireColumn.Width, _
                12.5)
                shp_ProjectRectangle.TextFrame.Characters.Text = "  " & vrt_ProjectStart & "  " & vrt_ProjectName
                'Selection.ShapeRange.Item("Line 3176").Width = 362.25
                With shp_ProjectRectangle
                    .Fill.Visible = True
                    .Fill.Transparency = 0.8
                    .Fill.ForeColor.SchemeColor = 12
                    .line.Weight = 1.25
                    .line.DashStyle = 1
                    .line.Style = 1
                    .line.Transparency = 0#
                    .line.Visible = True
                    .line.ForeColor.SchemeColor = 12
                    .line.BackColor.RGB = RGB(255, 255, 255)
                End With
            End If
        Next y
    Next x
End Sub

Private Sub PageInstructions()
'Modification of cells in this worksheet will affect other worksheets.
    Dim shp_Instruction As Shape
    Dim shp_Key As Shape
    Set shp_Instruction = Shapes.AddTextbox(msoTextOrientationHorizontal, Cells(20, 7).Left, Cells(20, 7).Top, 350, 45)    ' Cells(7, 13).Width,  Cells(7, 13).Height)
    shp_Instruction.TextFrame.Characters.Text = Chr(10) & "Modification of cells in this worksheet will affect other worksheets."
    With shp_Instruction
        .Fill.Visible = True
        .Fill.Transparency = 0.2
        .Fill.ForeColor.RGB = RGB(139, 137, 137)
    End With
End Sub
'http://support.microsoft.com/kb/213367
'http://www.codetoad.com/vba_watermark.asp
'expression.AddTextbox(Orientation, Left, Top, Width, Height)
'expression.Offset(RowOffset, ColumnOffset)
'expression.AddLine(BeginX, BeginY, EndX, EndY)



