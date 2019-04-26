Attribute VB_Name = "MainModule"
Sub ScheduleHelper()
    ScheduleHelperWindow.Show
End Sub

Function CreateNewFormat(isBold As Boolean, isItalic As Boolean, fontName As String, fontSize As String, fontColor As Long, cellColor As Long) As FormatStruct
    Set CreateNewFormat = New FormatStruct
    
    CreateNewFormat.vIsBold = isBold
    CreateNewFormat.vIsItalic = isItalic
    
    If fontName <> "" Then
        CreateNewFormat.vFontName = fontName
    Else
        CreateNewFormat.vFontName = "Calibri"
    End If
    
    If IsNumeric(fontSize) Then
        CreateNewFormat.vFontSize = fontSize
    Else
        CreateNewFormat.vFontSize = 34
    End If
    
    CreateNewFormat.vFontColor = fontColor
    CreateNewFormat.vCellColor = cellColor
End Function

Function FormatCell(YofCell As Integer, XofCell As Integer, inputFormat As FormatStruct)
    With ActiveWindow.Selection.ShapeRange(1).Table.Cell(YofCell, XofCell).Shape
        With .TextFrame.TextRange.Font
            .Bold = inputFormat.vIsBold
            .Italic = inputFormat.vIsItalic
            .size = inputFormat.vFontSize
            .Name = inputFormat.vFontName
            .Color.RGB = inputFormat.vFontColor
        End With
        With .Fill.ForeColor
            .RGB = inputFormat.vCellColor
        End With
    End With
End Function

Function NewSessionTime(n As Integer) As String
    'Determine if AM or PM
    Dim DayPeriod As String: DayPeriod = "AM"
    Dim PrevHour As Integer: PrevHour = 1
    Dim CurrHour As Integer
    With ActiveWindow.Selection.ShapeRange(1).Table
        For i = 2 To n
            With .Cell(i, 1).Shape.TextFrame.TextRange
                CurrHour = CInt(Split(.Text, ":")(0))
                If CurrHour = 12 Or CurrHour < PrevHour Then
                    DayPeriod = "PM"
                    GoTo LineReturn:
                End If
                PrevHour = CurrHour
            End With
        Next i
LineReturn:
        NewSessionTime = Split(.Cell(n, 1).Shape.TextFrame.TextRange.Text, " - ")(0) + DayPeriod
    End With
End Function


Function SessionTime(n As Integer) As String
    Dim TimeInt As Integer
    Dim Time As String
    
    If n <= 10 Then
        TimeInt = Int(8 + (n - 1) / 2)
        Time = "" + CStr(TimeInt)
        If (n Mod 2) <> 1 Then
            Time = Time + ":30AM"
        Else
            Time = Time + ":00AM"
        End If
    Else
        TimeInt = Int((n - 9) / 2)
        Time = "" + CStr(TimeInt)
        If (n Mod 2) <> 1 Then
            Time = Time + ":30PM"
        Else
            Time = Time + ":00PM"
        End If
    End If
    
    SessionTime = Time
End Function

Function InitializeBuiltInFormats() As FormatStruct()
    Dim retArr(1) As FormatStruct
    
    Dim bif1 As FormatStruct: Set bif1 = New FormatStruct
    With bif1
        .vDisplayName = "Normal class"
        .vCellColor = RGB(255, 153, 153)
        .vFontColor = RGB(0, 0, 0)
        .vFontName = "Gisha"
        .vFontSize = 34
        .vIsBold = False
        .vIsItalic = False
    End With
    Set retArr(0) = bif1
    
    Dim bif2 As FormatStruct: Set bif2 = New FormatStruct
    With bif2
        .vDisplayName = "Placement test"
        .vCellColor = RGB(255, 80, 80)
        .vFontColor = RGB(255, 255, 255)
        .vFontName = "Gisha"
        .vFontSize = 34
        .vIsBold = True
        .vIsItalic = False
    End With
    Set retArr(1) = bif2
    
    InitializeBuiltInFormats = retArr
End Function
