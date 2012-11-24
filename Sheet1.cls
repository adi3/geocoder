VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private InputFile
Private MyFileName As String
Private Legend As String
Private Region As String
Private ColorMin As String
Private ColorMax As String

'Selects dropdown option for map scope and shows/hides
'the display area selection options accordingly.

Sub ComboBox2_Change()
    If ComboBox2.Value = "World Map" Then
        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Label1.Visible = False
        Region = "world"
    ElseIf ComboBox2.Value = "Region Map" Then
        ComboBox3.Visible = True
        ComboBox4.Visible = False
        Label1.Visible = True
    ElseIf ComboBox2.Value = "Country Map" Then
        ComboBox3.Visible = False
        ComboBox4.Visible = True
        Label1.Visible = True
    End If
    
End Sub

'Dropdown options for selection region of the world on
'which map should be plotted. Visible only if ComboBox 2
'is set to "Region Map".

Sub ComboBox3_Change()
    Dim temp As Integer
    Dim reg As String
    temp = Application.WorksheetFunction.Match(ComboBox3.Value, Sheet2.Range("C1:C16"), 0)
    reg = "D" & temp
    Region = Sheet2.Range(reg).Value
    If Len(Region) = 1 Then
        Region = "00" & Region
    ElseIf Len(Region) = 2 Then
        Region = "0" & Region
    End If
    
End Sub

'Dropdown options for selection country on which map should
'be plotted. Visible only if ComboBox 2 is set to "Country Map".

Sub ComboBox4_Change()
    Dim temp As Integer
    Dim country As String
    temp = Application.WorksheetFunction.Match(ComboBox4.Value, Worksheets("Sheet2").Range("F1:F249"), 0)
    country = "G" & temp
    Region = Worksheets("Sheet2").Range(country).Value
    
End Sub

'Button to prompt user for selecting data file. Data files can
'only be an Excel file (.xls or .xlsx). Clicking the button opens
'up a dialog box that allows user to navigate through system
'folders and select a file.

Sub CommandButton2_Click()
    InputFile = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls), (*.xlsx; *.xls)")
    If InputFile <> False Then
        TextBox3.Value = InputFile
    End If
    
    Worksheets("Sheet2").Range("F252").Value = Module2.FileName(Worksheets("Sheet2").Range("F251"))
    ColorMax = Module2.Color(SpinButton1)
    ColorMin = Module2.Color(SpinButton2)

End Sub

'Sets SHOW Legend option to true, thereby displaying the legend
'boxes on the mockup image and the spin buttons for selecting
'heatmap colors.

Sub OptionButton3_Click()
    Legend = "true"
    Image2.Visible = True
    Image3.Visible = True
    SpinButton1.Visible = True
    SpinButton2.Visible = True
    Label2.Visible = True
    Label3.Visible = True
End Sub

'Sets HIDE Legend option to true, thereby hiding the legend
'boxes on the mockup image and the spin buttons for selecting
'heatmap colors.

Sub OptionButton4_Click()
    Legend = "false"
    Image2.Visible = False
    Image3.Visible = False
    SpinButton1.Visible = False
    SpinButton2.Visible = False
    Label2.Visible = False
    Label3.Visible = False
End Sub

'Horizontal scroll bar for changing mockup graph's width.
'Image1 (graph) changes width proportional to the scroll bar.
'Image2 and Image3 are legend boxes that scale accordingly too.

Sub ScrollBar1_Change()
    Image1.Width = ScrollBar1.Value / 7
    
    Image2.Width = Image1.Width / 8
    Image2.Left = Image1.Left + 3
    Image3.Width = Image1.Width / 8
    Image3.Left = Image1.Left + Image2.Width + 2
    
    Image2.Height = Image1.Height / 8
    Image2.Top = Image1.Top + Image1.Height - Image2.Height - 2
    Image3.Height = Image1.Height / 8
    Image3.Top = Image2.Top
End Sub

'Vertical scroll bar for changing mockup graph's height.
'Image1 (graph) changes height proportional to the scroll bar.
'Image2 and Image3 are legend boxes that scale accordingly too.

Sub ScrollBar2_Change()
    Image1.Height = ScrollBar2.Value / 5.5
    
    Image2.Width = Image1.Width / 8
    Image2.Left = Image1.Left + 3
    Image3.Width = Image1.Width / 8
    Image3.Left = Image1.Left + Image2.Width + 2
    
    Image2.Height = Image1.Height / 8
    Image2.Top = Image1.Top + Image1.Height - Image2.Height - 2
    Image3.Height = Image1.Height / 8
    Image3.Top = Image2.Top
End Sub

'Spin button for choosing maximum color of legend.
'Dynamically changes both its own color and the color
'of the maximum legend box in mockup about change.

Sub SpinButton1_Change()
    Worksheets("Sheet2").Range("E1").Interior.ColorIndex = Worksheets("Sheet2").Range("E2").Value
    SpinButton1.BackColor = Worksheets("Sheet2").Range("E1").Interior.Color
    Worksheets("Sheet2").Range("E1").Interior.ColorIndex = Null
    Image3.BackColor = SpinButton1.BackColor
    
    ColorMax = Module2.Color(SpinButton1)
    
End Sub

'Spin button for choosing minimum color of legend.
'Dynamically changes both its own color and the color
'of the minimum legend box in mockuup about change.

Sub SpinButton2_Change()
    Worksheets("Sheet2").Range("E3").Interior.ColorIndex = Worksheets("Sheet2").Range("E4").Value
    SpinButton2.BackColor = Worksheets("Sheet2").Range("E3").Interior.Color
    Worksheets("Sheet2").Range("E3").Interior.ColorIndex = Null
    Image2.BackColor = SpinButton2.BackColor
    
    ColorMin = Module2.Color(SpinButton2)
    
End Sub

'The generate plot button. Handles the creation of the
'HTML file in accordance to the options supplied by the UI.

Sub CommandButton1_Click()

    If TextBox3.Value = "" Then
        MsgBox "Please provide a data file path!", vbOKOnly, "Rolls-Royce | Adi - GeoCoder"
        Exit Sub
    End If
    
    Module3.AnimateProgress
    
    'Generates file. Returns file name for browser to open.
    MyFileName = Module3.GeneratePlot(Region, Legend, ColorMin, ColorMax)
    
    Module3.OpenBrowser (MyFileName)
    
End Sub
