Option Explicit

Private Sub Worksheet_Activate()
    Dim bln_Flag As Boolean
    bln_Flag = True
    Call AppEvents(bln_Flag)
End Sub

Private Sub Worksheet_Calculate()
    Call DynamicTextboxDisplay
    Call ProjectLengthDisplay
    Dim rng_Thing As Range
    Dim bln_Flag As Boolean
    Dim i As Integer
    Const xlColorIndexNone = -4142
    bln_Flag = False
    Call AppEvents(bln_Flag)
    Set rng_Thing = Range("M4:FI4")
    For Each rng_Thing In Range("M4:FI4") 'G4:AO4
        Select Case Application.Weekday(rng_Thing.Value)
            Case 1
                rng_Thing.Offset(1, 0).Value = "Su"
                rng_Thing.Offset(1, 0).Interior.Color = RGB(214, 214, 214)
            Case 2
                rng_Thing.Offset(1, 0).Value = "M"
                rng_Thing.Offset(1, 0).Interior.ColorIndex = xlColorIndexNone
            Case 3
                rng_Thing.Offset(1, 0).Value = "T"
                rng_Thing.Offset(1, 0).Interior.ColorIndex = xlColorIndexNone
            Case 4
                rng_Thing.Offset(1, 0).Value = "W"
                rng_Thing.Offset(1, 0).Interior.ColorIndex = xlColorIndexNone
            Case 5
                rng_Thing.Offset(1, 0).Value = "Th"
                rng_Thing.Offset(1, 0).Interior.ColorIndex = xlColorIndexNone
            Case 6
                rng_Thing.Offset(1, 0).Value = "F"
                rng_Thing.Offset(1, 0).Interior.ColorIndex = xlColorIndexNone
            Case 7
                rng_Thing.Offset(1, 0).Value = "Sa"
                rng_Thing.Offset(1, 0).Interior.Color = RGB(214, 214, 214)
            End Select
        Next
    Application.EnableCancelKey = xlInterrupt
    bln_Flag = True
    Call AppEvents(bln_Flag)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim bln_Flag As Boolean
    bln_Flag = True
    Call AppEvents(bln_Flag)
    Call DeleteTextboxes
    Call Worksheet_Calculate
    Call DynamicTextboxDisplay
End Sub

Private Sub AppEvents(bln_Flag)
'http://blogs.msdn.com/excel/archive/2009/03/12/excel-vba-performance-coding-best-practices.aspx
'the false bln_Flag turns off functionality
'the true bln_Flag turns it back on.
    Dim bln_ScreenUpdateState, bln_StatusBarState, bln_CalcState, bln_EventsState, bln_DisplayPageBreakState As Boolean
    'Get current state of various Excel settings - not used, but could be
    bln_ScreenUpdateState = Application.ScreenUpdating
    bln_StatusBarState = Application.DisplayStatusBar
    bln_CalcState = Application.Calculation
    bln_EventsState = Application.EnableEvents
    bln_DisplayPageBreakState = DisplayPageBreaks  'note this is a sheet-level setting
    If bln_Flag = False Then
        'turn off some Excel functionality so code runs faster
        With Application
            .ScreenUpdating = False
            .DisplayStatusBar = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
            '.DisplayPageBreaks = False 'note this is a sheet-level setting, can't set this property if you don't have a printer installed
            .CellDragAndDrop = False
            .CutCopyMode = False
            .EnableCancelKey = xlDisabled
        End With
    End If
    If bln_Flag = True Then
        'after code runs, restore state
        With Application
            .ScreenUpdating = True
            .DisplayStatusBar = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
            '.DisplayPageBreaks = True 'note this is a sheet-level setting, can't set this property if you don't have a printer installed
            .CellDragAndDrop = True
            .CutCopyMode = True
            .EnableCancelKey = xlDisabled
        End With
    End If
End Sub
Function DynamicTextboxDisplay()
    'http://msdn.microsoft.com/en-us/library/bb242246.aspx
    'display a textbox with project info, worksheet instruction, and the key
    Dim shp_ProjectDate As Shape
    Dim shp_Instruction As Shape
    Dim shp_Key As Shape
    'project date estimate
    Set shp_ProjectDate = Shapes.AddTextbox(msoTextOrientationHorizontal, Cells(7, 10).Left, Cells(2, 10).Top, 200, 45)    ' Cells(7, 13).Width,  Cells(7, 13).Height)
    shp_ProjectDate.TextFrame.Characters.Text = "Project Length (includes client follow-up): " & Chr(10) & Chr(10) & Cells(7, 5) & " to " & Cells(44, 6)    '"test"
    Call FormatDTD(shp_ProjectDate)
    'worksheet instruction
    Set shp_Instruction = Shapes.AddTextbox(msoTextOrientationHorizontal, Cells(22, 15).Left, Cells(24, 10).Top, 250, 120)    ' Cells(7, 13).Width,  Cells(7, 13).Height)
    shp_Instruction.TextFrame.Characters.Text = _
    "Instructions" & Chr(10) & Chr(10) & "1) Enter the project's Initial Review start date" & Chr(10) & Chr(10) & "2) Enter the % Complete for the sub-task only (Blue-Gray cells)" & Chr(10) & Chr(10) & "WARNING: Do NOT modify other cells as they contain formulas and are dependent on one another."
    Call FormatDTD(shp_Instruction)
    'gantt key
    Set shp_Key = Shapes.AddTextbox(msoTextOrientationHorizontal, Cells(22, 15).Left, Cells(34, 15).Top, 250, 40)    ' Cells(7, 13).Width,  Cells(7, 13).Height)
    shp_Key.TextFrame.Characters.Text = "Blue cells represent remaining days of project task." & Chr(10) & Chr(10) & "Green cells represent completed days of project task."
    Call FormatDTD(shp_Key)
End Function
Private Sub FormatDTD(ByRef boxName As Shape)
    'assign common formatting to textboxes
    With boxName
        .Fill.Transparency = 0.8
        .Fill.ForeColor.RGB = RGB(0, 51, 102)
        '.Fill.BackColor.RGB = RGB(255, 255, 255)
        '.Fill.TwoColorGradient msoGradientHorizontal, 1 'Vertical, 1
    End With
End Sub
Private Sub DeleteTextboxes()
    Dim i As Integer
    For i = Shapes.Count To 1 Step -1
         Shapes(i).Delete
    Next i
End Sub

Private Sub ProjectLengthDisplay()
    'look at each cell in the column and compare to each cell in the row
    'if dates match, display a rectangle over the task dates
    Call DeleteTextboxes
   Dim varMatch
   Dim lng_ProjectDuration As Long
   Dim shp_ProjectInfoDisplay As Excel.Shape, shp_MilestoneMarker As Excel.Shape, shp_ProjectRectangle As Excel.Shape
   Dim shp_ProjectItemInfo As Excel.Shape
   Dim vrt_ProjectName, vrt_ProjectStart, vrt_ProjectEnd As Variant
   Dim rng_CompareRow As Range, rng_CompareCol As Range
   Dim x As Range
   Dim str_MilestoneInfo As String
   str_MilestoneInfo = "Milestone 1 & 2" & Chr(10) & "Start" & Chr(10) & "Design Approval"
   Set rng_CompareRow = Range("M4:FI4")
   Set rng_CompareCol = Range("E7:E44")
   For Each x In rng_CompareCol
      varMatch = Application.Match(x.Value2, rng_CompareRow, 0)
      If Not IsError(varMatch) Then
         vrt_ProjectName = CStr(x.Offset(0, -2).Value)
         vrt_ProjectStart = x.Value
         vrt_ProjectEnd = x.Offset(0, 1).Value
         lng_ProjectDuration = (x.Offset(0, 4).Value)
         With Range("M:FI").Cells(x.Row, varMatch)
            Set shp_ProjectRectangle = _
             Shapes.AddShape( _
            msoShapeRectangle, .Left, .Top, _
            lng_ProjectDuration * .EntireColumn.Width, _
            12.5)
            'add a textbox offset by 0 rows and -1 cols
            Set shp_ProjectItemInfo = _
             Shapes.AddTextbox(msoTextOrientationHorizontal, _
            .Offset(1, 0).Left, .Offset(1, 0).Top, _
            150, 23)
         End With
         With shp_ProjectRectangle
            With .Fill
               .Visible = True
               .Transparency = 0.9
               '.ForeColor.SchemeColor = 52
               .ForeColor.RGB = RGB(255, 255, 255)
            End With
            With .line
               .Weight = 1.25
               .DashStyle = 1
               .Style = 1
               .Transparency = 0#
               .Visible = True
               '.ForeColor.SchemeColor = 2
               .ForeColor.RGB = RGB(0, 51, 102)
               .BackColor.RGB = RGB(255, 255, 255)
            End With
         End With
         With shp_ProjectItemInfo
            With .TextFrame
                .Characters.Font.Size = 9
                .Characters.Font.Bold = True
                .Characters.Text = " " & vrt_ProjectStart & "  " & vrt_ProjectName & Chr(10) & " " & vrt_ProjectEnd
            End With
            With .Fill
                .Visible = True
                .Transparency = 0.7
                '.ForeColor.RGB = RGB(255, 255, 255)
                .ForeColor.RGB = RGB(0, 51, 102)
            End With
            With .line
               .Weight = 1.25
               .DashStyle = 1
               .Style = 1
               .Transparency = 0#
               .Visible = True
               '.ForeColor.SchemeColor = 2
               .ForeColor.RGB = RGB(0, 51, 102)
               .BackColor.RGB = RGB(255, 255, 255)
            End With
         End With
      End If
   Next x
End Sub

'http://support.microsoft.com/kb/213367
'http://www.codetoad.com/vba_watermark.asp
'expression.AddTextbox(Orientation, Left, Top, Width, Height)
'expression.Offset(RowOffset, ColumnOffset)
'expression.AddLine(BeginX, BeginY, EndX, EndY)
'http://cloford.com/resources/colours/500col.htm
'http://web.njit.edu/~kevin/rgb.txt.html
'http://www.tayloredmktg.com/rgb/#BL

Private Sub xlVariables()
   MsgBox xlNone ' -4142
   MsgBox xlContinuous ' 1
   MsgBox xlMedium ' -4138
   MsgBox xlAutomatic ' -4105
   MsgBox xlDiagonalDown ' 5
   MsgBox xlDiagonalUp ' 6
   MsgBox xlEdgeLeft ' 7
   MsgBox xlEdgeTop ' 8
   MsgBox xlEdgeBottom ' 9
   MsgBox xlEdgeRight ' 10
   MsgBox xlInsideVertical ' 11
   MsgBox xlInsideHorizontal ' 12
   MsgBox xlDistributed ' -4117
   MsgBox xlCenter ' -4108
End Sub
