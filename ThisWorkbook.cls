VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
 
Sub Application_Open()

    'Changes app icon as defined in Module 1
    ChangeApplicationIcon
    
End Sub

Sub Workbook_Activate()
    Application.Caption = "Rolls-Royce | Adi"
    
    'Called again as icon is changed by Excel each time a workbook activates
    ChangeApplicationIcon
    
    'The progress bar animation shown upon clicking Generate Plot.
    Sheet1.Shapes("Rectangle 59").Visible = msoFalse
    
    Sheet1.ComboBox2.Value = "World Map"
    
    If Sheet1.TextBox1.Value <> "" Then
        Sheet1.TextBox1.Value = ""
    End If
    If Sheet1.TextBox2.Value <> "" Then
        Sheet1.TextBox2.Value = ""
    End If
    If Sheet1.TextBox3.Value <> "" Then
        Sheet1.TextBox3.Value = ""
    End If
    
    ActiveWindow.DisplayHeadings = True

    On Error Resume Next

        With Application
        
           .DisplayFullScreen = True

           .CommandBars("Standard").Visible = False
           
           .CommandBars("Formatting").Visible = False
           
           .CommandBars("Drawing").Visible = False
           
           .CommandBars("Status Bar").Visible = False
           
           .CommandBars("Control Toolbox").Visible = False
                      
           .CommandBars("Formula Bar").Visible = False
           
           .CommandBars("Status Bar").Visible = False

           .CommandBars("Worksheet Menu Bar").Enabled = False

           .DisplayStatusBar = False
           
           .DisplayFormulaBar = False
           
           .Cursor = xlNorthwestArrow

        End With
        
    On Error GoTo 0

    'Hides row and column headers
    ActiveWindow.DisplayHeadings = False

End Sub

Sub Workbook_BeforeClose(Cancel As Boolean)

    'To prevent program from prompting user to save work before exit.
    Me.Saved = True

End Sub

Sub Workbook_Deactivate()
    
    'Excel settings seem to persist between programs.
    'So it is necessary to reset them before exiting.
    
    Application.Caption = "Microsoft Excel"
    ChangeApplicationIcon
    
    ActiveWindow.DisplayHeadings = True

     On Error Resume Next

        With Application

           .DisplayFullScreen = False
                       
           .CommandBars("Standard").Visible = True
           
           .CommandBars("Formatting").Visible = True
           
           .CommandBars("Drawing").Visible = True
           
           .CommandBars("Status Bar").Visible = True
           
           .CommandBars("Control Toolbox").Visible = True
                      
           .CommandBars("Formula Bar").Visible = True
           
           .CommandBars("Status Bar").Visible = True

           .CommandBars("Worksheet Menu Bar").Enabled = True

           .DisplayStatusBar = True
           
           .DisplayFormulaBar = True
           
           .Cursor = xlDefault
           
        End With

    On Error GoTo 0


End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    MsgBox "adi"
End Sub
