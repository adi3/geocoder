Attribute VB_Name = "Module3"
Option Explicit

'Smoothly fills a rectangular bar at the bottom of form to give
'an impression of data processing.
Sub AnimateProgress()

    Worksheets("Sheet1").Shapes("Rectangle 59").Visible = msoTrue
    Worksheets("Sheet1").Shapes("Rectangle 59").Width = 0
    Dim Completed As Single
    Dim A As Integer
    Dim B As Integer
    
    For B = 1 To 650
        For A = 1 To 10000
            Next
                DoEvents
                Worksheets("Sheet1").Shapes("Rectangle 59").Width = B
            Next
                MsgBox "Success: File Created and Exported", vbOKOnly, "Rolls-Royce | Adi - GeoCoder"
       
End Sub

'Generates the HTML file with Google Chart API references
'and returns the file path.
Function GeneratePlot(Region As String, Legend As String, ColorMin As String, ColorMax As String)
      
    Dim Mode As String
    Dim ExcelPath As String
    Dim objExcel As Object
    Dim MyFileName As String
    
    Dim Active
    Dim CurrentWorkSheet
    Dim UsedRowsCount
    Dim Row
    Dim Column
    Dim Top
    Dim Left
    Dim Cells
    Dim CurCol
    Dim CurRow
    Dim Word
    
    ExcelPath = Worksheets("Sheet1").TextBox3.Value
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Workbooks.Open ExcelPath, False, True
    Set CurrentWorkSheet = objExcel.ActiveWorkbook.Worksheets(1)
    Set Cells = CurrentWorkSheet.Cells
    
    UsedRowsCount = CurrentWorkSheet.UsedRange.Rows.Count
    Top = CurrentWorkSheet.UsedRange.Row
    Left = CurrentWorkSheet.UsedRange.Column
    
    If Worksheets("Sheet1").OptionButton1.Value = True Then
        Mode = "markers"
    Else
        Mode = "regions"
    End If
    
    If Region = "" Then
        Region = "world"
    End If
    
    If Legend = "" Then
        Legend = "true"
    End If
    
    If ColorMax = "" Then
        ColorMax = "008000"
    End If
    If ColorMin = "" Then
        ColorMin = "800000"
    End If
        
    MyFileName = Worksheets("Sheet2").Range("F253").Value & Worksheets("Sheet1").TextBox1.Value & ".html"
    Open MyFileName For Output As #1
     
    Print #1, "<html>"
    Print #1, "<head>"
    Print #1, "<title>Rolls-Royce | Adi - GeoCoder - " & Worksheets("Sheet1").TextBox1.Value & "</title>"
    Print #1, "<script type='text/javascript' src='https://www.google.com/jsapi'></script>"
    Print #1, "<script type='text/javascript'>"
    Print #1, ""
    Print #1, "    google.load('visualization', '1', {'packages': ['geomap']});"
    Print #1, "    google.setOnLoadCallback(drawMap);"
    Print #1, ""
    Print #1, "    function drawMap() {"
    Print #1, "        var data = new google.visualization.DataTable();"

    Print #1, "        data.addRows(" & UsedRowsCount & ");"
    Print #1, "        data.addColumn('string', '" & Worksheets("Sheet1").ComboBox1.Value & "');"
    Print #1, "        data.addColumn('number', '" & Worksheets("Sheet1").TextBox2.Value & "');"

    For Row = 0 To (UsedRowsCount - 1)
        
        For Column = 0 To 1
            CurRow = Row + Top
            CurCol = Column + Left
            Word = Cells(CurRow, CurCol).Value
            
           If Word <> "" Then
                If CurCol Mod 2 Then
                    Print #1, "        data.setValue(" & CurRow - 1 & "," & CurCol - 1 & ",'" & Word & "');"
                Else
                    Print #1, "        data.setValue(" & CurRow - 1 & "," & CurCol - 1 & "," & Word & ");"
                End If
           End If
            
        Next
    Next
    
    Print #1, ""
    Print #1, "     var options = {};"
    Print #1, "     options['region'] = '" & Region & "';"
    Print #1, "     options['dataMode'] = '" & Mode & "';"
    Print #1, "     options['width'] = '" & Worksheets("Sheet1").ScrollBar1.Value & "px';"
    Print #1, "     options['height'] = '" & Worksheets("Sheet1").ScrollBar2.Value & "px';"
    
    Print #1, "     options['colors'] = [0x" & ColorMin & ", 0x" & ColorMax & "];"
    Print #1, "     options['showLegend'] = " & Legend & ";"
    
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
    GeneratePlot = MyFileName
    
End Function

'Opens the received file in an IE instance.
Sub OpenBrowser(FileName As String)

    Dim ie As Object
    Set ie = CreateObject("Internetexplorer.Application")
    ie.Visible = True
    ie.Navigate FileName
    
End Sub

