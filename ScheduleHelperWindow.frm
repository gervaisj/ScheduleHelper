VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScheduleHelperWindow 
   Caption         =   "Schedule Helper"
   ClientHeight    =   4280
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6660
   OleObjectBlob   =   "ScheduleHelperWindow.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScheduleHelperWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The built-in formats available in the Format Preset drop-down list
' Others can be added in the InitializeBuiltInFormats function in MainModule
Dim BuiltInFormats() As FormatStruct

'Format currently used for the preview and when adding a class
Dim CurrentFormat As FormatStruct

'Some booleans to make sure the _Change events don't get triggered during initialization of the userform
Dim FormatPresetComboBox_initialized As Boolean
Dim SizeComboBox_initialized As Boolean

Private Sub UserForm_Initialize()
    'Dim ds As Date: ds = DateSerial(2019, 4, 27)
    'ds = DateAdd("d", 1, ds)
    'Dim str As String: str = WeekdayName(Weekday(ds)) & " " & FormatDateTime(ds, vbLongDate)

    'Populate the size combobox
    SizeComboBox_initialized = False
    Dim size As Integer
    For size = 20 To 60
        SizeComboBox.AddItem (CStr(size))
    Next
    SizeComboBox_initialized = True
    
    'Init. the built-in formats and set the first one as the current one
    BuiltInFormats = InitializeBuiltInFormats
    Set CurrentFormat = BuiltInFormats(0)
    
    'Populate the format presets combobox using the array BuiltInFormats
    FormatPresetComboBox_initialized = False
    With FormatPresetComboBox
        For i = 0 To UBound(BuiltInFormats)
            .AddItem (i)
            .Column(1, i) = BuiltInFormats(i).vDisplayName
        Next i
        .Value = 0
    End With
    FormatPresetComboBox_initialized = True
End Sub

Private Sub AddButton_Click()

    'Bogus operation to pad the Undo stack so that undoing won't delete all the previously added classes
    With ActivePresentation.Slides(1).Background.Fill.BackColor
        .RGB = .RGB
    End With

    Dim TLCell_X As Integer 'x coordinate of the top left selected cell
    Dim TLCell_Y As Integer 'y coordinate...
    Dim BRCell_X As Integer 'x coordinate of the bottom right selected cell
    Dim BRCell_Y As Integer 'y coordinate...
    
    Dim Class As String: Class = ClassTextBox.Text
    Dim Instructor As String: Instructor = InstructorTextBox.Text
    
    Dim StartTime As String 'start time of the class
    Dim EndTime As String 'end time...
    Dim OutputToCell As String 'Final output to the merged cell in the schedule
    
    Call FindTLCell(TLCell_Y, TLCell_X) 'pass by ref
    
    Call FindBRCell(BRCell_Y, BRCell_X) 'pass by ref
    
    'Determine the start time of the selection
    StartTime = SessionTime(TLCell_Y)
    
    'Determine the end time of the selection
    EndTime = SessionTime(BRCell_Y + 1)
    
    'Check if the range selected is valid and set the format of the output
    If TLCell_X = BRCell_X And TLCell_Y = BRCell_Y Then
        MsgBox "Please select more than one cell"
        Exit Sub
    ElseIf BRCell_Y - TLCell_Y = 1 Then
        'Height = 2
        OutputToCell = Class + " | " + Instructor + vbNewLine + StartTime + " - " + EndTime
    ElseIf BRCell_Y - TLCell_Y >= 2 Then
        'Height > 2
        OutputToCell = Class + vbNewLine + Instructor + vbNewLine + StartTime + " - " + EndTime
    Else
        'Height = 1, width > 1
        OutputToCell = Class + " | " + Instructor + " | " + StartTime + " - " + EndTime
    End If
    
    With ActiveWindow.Selection.ShapeRange(1).Table
        'Merge cells that are selected
        .Cell(TLCell_Y, TLCell_X).Merge MergeTo:=.Cell(BRCell_Y, BRCell_X)
        
        With .Cell(TLCell_Y, TLCell_X).Shape.TextFrame
            'Paste text to cell
            .TextRange.Text = OutputToCell
            
            'Center the text on the vertical axis
            .VerticalAnchor = msoAnchorMiddle
            
            'Center the text on the horizontal axis
            .HorizontalAnchor = msoAnchorCenter
        End With
    End With
        
    Call FormatCell(TLCell_Y, TLCell_X, CurrentFormat)
    
    'Clear the fields of the user form if the option is selected
    If ClearFieldsCheckBox.Value = True Then
        ClassTextBox.Text = ""
        InstructorTextBox.Text = ""
    End If
    
End Sub

Private Sub EnforceFormatCheckBox_Click()
    'If the Enforce format checkbox is ticked, enable the relevant controls and update the preview
    If EnforceFormatCheckBox.Value = True Then
        Call EnableCustomFormatControls
        Call UpdatePreviewHandler
    'Do the opposite otherwise, while still updating the preview
    Else
        Call DisableCustomFormatControls
        Call UpdatePreviewHandler
    End If
End Sub

Private Function UpdatePreviewHandler()
    'Use the selected built-in format as the current format to update the preview
    If EnforceFormatCheckBox.Value = False Then
        Set CurrentFormat = BuiltInFormats(CInt(FormatPresetComboBox.Value))
    'Otherwise, create a new one using the "custom format" controls
    Else
        Set CurrentFormat = CreateNewFormatFromForm
    End If
    'Update the preview using CurrentFormat
    Call UpdatePreview(CurrentFormat)
End Function

Private Function UpdatePreview(SelectedFormatPreset As FormatStruct)
    'Changing the properties...
    PreviewText.BackColor = SelectedFormatPreset.vCellColor
    PreviewText.ForeColor = SelectedFormatPreset.vFontColor
    With PreviewText.Font
        .Bold = SelectedFormatPreset.vIsBold
        .Italic = SelectedFormatPreset.vIsItalic
        .Name = SelectedFormatPreset.vFontName
        .size = SelectedFormatPreset.vFontSize
    End With
End Function

Private Function CreateNewFormatFromForm() As FormatStruct
    On Error GoTo ReturnCurrentFormat:
        'Try to convert the text in the color textboxes to integers
        Dim Font_R As Integer: Font_R = CInt(FontColor_RedTextBox.Text)
        Dim Font_G As Integer: Font_G = CInt(FontColor_GreenTextBox.Text)
        Dim Font_B As Integer: Font_B = CInt(FontColor_BlueTextBox.Text)
        Dim Cell_R As Integer: Cell_R = CInt(CellColor_RedTextBox.Text)
        Dim Cell_G As Integer: Cell_G = CInt(CellColor_GreenTextBox.Text)
        Dim Cell_B As Integer: Cell_B = CInt(CellColor_BlueTextBox.Text)
        'Note that RGB() can also throw an error if the values are negative
        Set CreateNewFormatFromForm = CreateNewFormat(BoldCheckBox.Value, ItalicCheckBox.Value, _
                                                        FontTextBox.Text, SizeComboBox.Text, _
                                                        RGB(Font_R, Font_G, Font_B), _
                                                        RGB(Cell_R, Cell_G, Cell_B))
        Exit Function
        'If CInt failed to convert and raised an exception, return the current format (i.e. don't change anything)
ReturnCurrentFormat:
        Set CreateNewFormatFromForm = CurrentFormat
        
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' After this point, it's a bunch of event handler definitions
' And utility methods that hide tedious code
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub BoldCheckBox_Click()
    Call UpdatePreviewHandler
End Sub

Private Sub CellColor_BlueTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub CellColor_GreenTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub CellColor_RedTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub FontColor_BlueTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub FontColor_GreenTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub FontColor_RedTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub FontTextBox_Change()
    Call UpdatePreviewHandler
End Sub

Private Sub ItalicCheckBox_Click()
    Call UpdatePreviewHandler
End Sub

Private Sub FormatPresetComboBox_Change()
    If FormatPresetComboBox_initialized = False Then
        Exit Sub
    End If
    Call UpdatePreviewHandler
End Sub

Private Sub SizeComboBox_Change()
    If SizeComboBox_initialized = False Then
        Exit Sub
    End If
    Call UpdatePreviewHandler
End Sub

Private Function FindTLCell(ByRef TL_Y As Integer, ByRef TL_X As Integer)
    With ActiveWindow.Selection.ShapeRange(1).Table
        'Find the Top-Left cell (TLCell) of the selection
        For X = 1 To .Columns.Count
            For Y = 1 To .Rows.Count
                If .Cell(Y, X).Selected <> False Then
                    TL_X = X
                    TL_Y = Y
                    Exit Function
                End If
            Next
        Next
    End With
End Function

Private Function FindBRCell(ByRef BR_Y As Integer, ByRef BR_X As Integer)
    With ActiveWindow.Selection.ShapeRange(1).Table
        'Find the Bottom-Right cell (BRCell) of the selection
        For X = .Columns.Count To 1 Step -1
            For Y = .Rows.Count To 1 Step -1
                If .Cell(Y, X).Selected <> False Then
                    BR_X = X
                    BR_Y = Y
                    Exit Function
                End If
            Next
        Next
    End With
End Function

Private Function EnableCustomFormatControls()
    BoldCheckBox.Enabled = True
    ItalicCheckBox.Enabled = True
    SizeComboBox.Enabled = True
    SizeComboBox.BackColor = &H80000005
    FontLabel.Enabled = True
    FontTextBox.Enabled = True
    FontTextBox.BackColor = &H80000005
    FontColorFrame.Enabled = True
    FontColor_RedLabel.Enabled = True
    FontColor_GreenLabel.Enabled = True
    FontColor_BlueLabel.Enabled = True
    FontColor_RedTextBox.Enabled = True
    FontColor_GreenTextBox.Enabled = True
    FontColor_BlueTextBox.Enabled = True
    
    FontColor_RedTextBox.BackColor = &H80000005
    FontColor_GreenTextBox.BackColor = &H80000005
    FontColor_BlueTextBox.BackColor = &H80000005
    
    CellColorFrame.Enabled = True
    
    CellColor_RedLabel.Enabled = True
    CellColor_GreenLabel.Enabled = True
    CellColor_BlueLabel.Enabled = True
    
    CellColor_RedTextBox.Enabled = True
    CellColor_GreenTextBox.Enabled = True
    CellColor_BlueTextBox.Enabled = True
    
    CellColor_RedTextBox.BackColor = &H80000005
    CellColor_GreenTextBox.BackColor = &H80000005
    CellColor_BlueTextBox.BackColor = &H80000005
    
    FormatPresetLabel.Enabled = False
    FormatPresetComboBox.Enabled = False
    FormatPresetComboBox.BackColor = &H80000004
End Function

Private Function DisableCustomFormatControls()
    BoldCheckBox.Enabled = False
    ItalicCheckBox.Enabled = False
    SizeComboBox.Enabled = False
    SizeComboBox.BackColor = &H80000004
    FontLabel.Enabled = False
    FontTextBox.Enabled = False
    FontTextBox.BackColor = &H80000004
    
    FontColorFrame.Enabled = False
    
    FontColor_RedLabel.Enabled = False
    FontColor_GreenLabel.Enabled = False
    FontColor_BlueLabel.Enabled = False
    
    FontColor_RedTextBox.Enabled = False
    FontColor_GreenTextBox.Enabled = False
    FontColor_BlueTextBox.Enabled = False
    
    FontColor_RedTextBox.BackColor = &H80000004
    FontColor_GreenTextBox.BackColor = &H80000004
    FontColor_BlueTextBox.BackColor = &H80000004
    
    CellColorFrame.Enabled = False
    
    CellColor_RedLabel.Enabled = False
    CellColor_GreenLabel.Enabled = False
    CellColor_BlueLabel.Enabled = False
    
    CellColor_RedTextBox.Enabled = False
    CellColor_GreenTextBox.Enabled = False
    CellColor_BlueTextBox.Enabled = False
    
    CellColor_RedTextBox.BackColor = &H80000004
    CellColor_GreenTextBox.BackColor = &H80000004
    CellColor_BlueTextBox.BackColor = &H80000004
    
    FormatPresetLabel.Enabled = True
    FormatPresetComboBox.Enabled = True
    FormatPresetComboBox.BackColor = &H80000005
End Function
