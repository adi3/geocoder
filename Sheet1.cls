VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public legend As String
Public region As String
Public colormin As String
Public colormax As String

Sub ComboBox2_Change()
    If ComboBox2.Value = "World Map" Then
        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Label1.Visible = False
        region = "world"
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

Sub ComboBox3_Change()
Dim temp As Integer
Dim reg As String
    temp = Application.WorksheetFunction.Match(ComboBox3.Value, Worksheets("Sheet2").Range("C1:C16"), 0)
    reg = "D" & temp
    region = Worksheets("Sheet2").Range(reg).Value
        If Len(region) = 1 Then
            region = "00" & region
        ElseIf Len(region) = 2 Then
            region = "0" & region
        End If
End Sub

Sub ComboBox4_Change()
Dim temp As Integer
Dim country As String
    temp = Application.WorksheetFunction.Match(ComboBox4.Value, Worksheets("Sheet2").Range("F1:F249"), 0)
    country = "G" & temp
    region = Worksheets("Sheet2").Range(country).Value
End Sub

Sub CommandButton1_Click()
    
Dim Value As String
Dim mode As String

Dim Active
Dim currentWorkSheet
Dim usedRowsCount
Dim row
Dim column
Dim top
Dim left
Dim Cells
Dim curCol
Dim curRow
Dim word


MyFileName = Worksheets("Sheet2").Range("F253").Value & TextBox1.Value & ".html"

 Open MyFileName For Output As #1
 
 Print #1, "<html>"
 Print #1, "<head>"
 Print #1, "<title>Rolls-Royce | Adi - GeoCoder - " & TextBox1.Value & "</title>"
 Print #1, "<script type='text/javascript' src='https://www.google.com/jsapi'></script>"
 Print #1, "<script type='text/javascript'>"
 Print #1, ""
 Print #1, "    google.load('visualization', '1', {'packages': ['geomap']});"
 Print #1, "    google.setOnLoadCallback(drawMap);"
 Print #1, ""
 Print #1, "    function drawMap() {"
 Print #1, "        var data = new google.visualization.DataTable();"
 
    excelPath = TextBox3.Value
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Workbooks.Open excelPath, False, True
    Set currentWorkSheet = objExcel.ActiveWorkbook.Worksheets(1)

    usedRowsCount = currentWorkSheet.UsedRange.Rows.Count

 Print #1, "        data.addRows(" & usedRowsCount & ");"
 Print #1, "        data.addColumn('string', '" & ComboBox1.Value & "');"
 Print #1, "        data.addColumn('number', '" & TextBox2.Value & "');"
    
    top = currentWorkSheet.UsedRange.row
    left = currentWorkSheet.UsedRange.column

    Set Cells = currentWorkSheet.Cells
    
    For row = 0 To (usedRowsCount - 1)
        
        For column = 0 To 1
            curRow = row + top
            curCol = column + left
            word = Cells(curRow, curCol).Value
            
           If word <> "" Then
                If curCol Mod 2 Then
                    Print #1, "        data.setValue(" & curRow - 1 & "," & curCol - 1 & ",'" & word & "');"
                Else
                    Print #1, "        data.setValue(" & curRow - 1 & "," & curCol - 1 & "," & word & ");"
                End If
           End If
            
        Next
    Next

        If OptionButton1.Value = True Then
            mode = "markers"
        Else
            mode = "regions"
        End If

Print #1, ""
Print #1, "     var options = {};"
    If region = "" Then
        region = "035"
    End If
Print #1, "     options['region'] = '" & region & "';"
Print #1, "     options['dataMode'] = '" & mode & "';"
Print #1, "     options['width'] = '" & ScrollBar1.Value & "px';"
Print #1, "     options['height'] = '" & ScrollBar2.Value & "px';"

    If colormax = "" Then
        colormax = "008000"
    End If
    If colormin = "" Then
        colormin = "800000"
    End If
Print #1, "     options['colors'] = [0x" & colormin & ", 0x" & colormax & "];"
    If legend = "" Then
        legend = "true"
    End If
Print #1, "     options['showLegend'] = " & legend & ";"

Print #1, ""
Print #1, "     var container = document.getElementById('map_canvas');"
Print #1, "     var geomap = new google.visualization.GeoMap(container);"
Print #1, "     geomap.draw(data, options);"
Print #1, "};"
Print #1, ""
Print #1, "</script>"
Print #1, "</head>"

Print #1, ""
Print #1, "<body>"
Print #1, "     <div id='map_canvas'></div>"
Print #1, "</body>"
Print #1, "</html>"

Close #1

Shapes("Rectangle 59").Visible = msoTrue
Shapes("Rectangle 59").Width = 0
   Dim Completed As Single
   For b = 1 To 650
   For a = 1 To 60000
    a = a
   Next
        DoEvents
        Shapes("Rectangle 59").Width = b
    Next
    
MsgBox "Success: File Created and Exported", vbOKOnly, "Rolls-Royce | Adi - GeoCoder"

    Dim ie As Object
    Set ie = CreateObject("Internetexplorer.Application")
    ie.Visible = True
    ie.Navigate MyFileName
    
End Sub

Sub CommandButton2_Click()
    FilePath = Application.GetOpenFilename("Excel Files (*.xls), *.xls")
    If FilePath <> False Then
        TextBox3.Value = FilePath
    End If
    
    Dim i As Integer
    Dim Pos As Integer
    Dim FindChar As String
    Dim SearchString As String
 
    SearchString = Worksheets("Sheet2").Range("F251")
    FindChar = "\"
 
    For i = 1 To Len(SearchString)
        If Mid(SearchString, i, 1) = FindChar Then
            Pos = i
        End If
    Next i
    Worksheets("Sheet2").Range("F252").Value = Pos

End Sub

Sub OptionButton3_Click()
    legend = "true"
    Image2.Visible = True
    Image3.Visible = True
    SpinButton1.Visible = True
    SpinButton2.Visible = True
    Label2.Visible = True
    Label3.Visible = True
End Sub

Sub OptionButton4_Click()
    legend = "false"
    Image2.Visible = False
    Image3.Visible = False
    SpinButton1.Visible = False
    SpinButton2.Visible = False
    Label2.Visible = False
    Label3.Visible = False
End Sub

Sub ScrollBar1_Change()
    Image1.Width = ScrollBar1.Value / 7
    
    Image2.Width = Image1.Width / 8
    Image2.left = Image1.left + 3
    Image3.Width = Image1.Width / 8
    Image3.left = Image1.left + Image2.Width + 2
    
    Image2.Height = Image1.Height / 8
    Image2.top = Image1.top + Image1.Height - Image2.Height - 2
    Image3.Height = Image1.Height / 8
    Image3.top = Image2.top
End Sub

Sub ScrollBar2_Change()
    Image1.Height = ScrollBar2.Value / 5.5
    
    Image2.Width = Image1.Width / 8
    Image2.left = Image1.left + 3
    Image3.Width = Image1.Width / 8
    Image3.left = Image1.left + Image2.Width + 2
    
    Image2.Height = Image1.Height / 8
    Image2.top = Image1.top + Image1.Height - Image2.Height - 2
    Image3.Height = Image1.Height / 8
    Image3.top = Image2.top
End Sub

Sub SpinButton1_Change()
    Worksheets("Sheet2").Range("E1").Interior.ColorIndex = Worksheets("Sheet2").Range("E2").Value
    SpinButton1.BackColor = Worksheets("Sheet2").Range("E1").Interior.color
    Worksheets("Sheet2").Range("E1").Interior.ColorIndex = Null
    Image3.BackColor = SpinButton1.BackColor
    
    Dim temp As Integer
    Dim col As String
    temp = Application.WorksheetFunction.Match(SpinButton1.Value, Worksheets("Sheet2").Range("H1:H56"), 0)
    col = "J" & temp
    colormax = Worksheets("Sheet2").Range(col).Value
End Sub


Sub SpinButton2_Change()
    Worksheets("Sheet2").Range("E3").Interior.ColorIndex = Worksheets("Sheet2").Range("E4").Value
    SpinButton2.BackColor = Worksheets("Sheet2").Range("E3").Interior.color
    Worksheets("Sheet2").Range("E3").Interior.ColorIndex = Null
    Image2.BackColor = SpinButton2.BackColor
    
    Dim temp As Integer
    Dim col As String
    temp = Application.WorksheetFunction.Match(SpinButton2.Value, Worksheets("Sheet2").Range("H1:H56"), 0)
    col = "J" & temp
    colormin = Worksheets("Sheet2").Range(col).Value
End Sub
