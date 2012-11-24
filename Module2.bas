Attribute VB_Name = "Module2"
Option Explicit

'Parses and returns file name from the received string value
'so as to save the output file with the same name as input
'file but .html extension (if no name is supplied by user).

 Function FileName(Param As Range)
 
    Dim i As Integer
    Dim Pos As Integer
    Dim FindChar As String
    Dim SearchString As String

    SearchString = Param
    FindChar = "\"

    For i = 1 To Len(SearchString)
        If Mid(SearchString, i, 1) = FindChar Then
            Pos = i
        End If
    Next i
    FileName = Pos
    
 End Function
 
 'Returns the hex color code of the received SpinButton by checking it
 'against a previously stored map in Sheet 2.
 
  Function Color(Param As SpinButton)
  
    Dim col As String
    col = "J" & Param.Value
    Color = Sheet2.Range(col).Value
    
 End Function
