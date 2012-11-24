Attribute VB_Name = "Module1"
Option Explicit
 
Declare Function GetActiveWindow32 Lib "USER32" Alias _
"GetActiveWindow" () As Integer
 
Declare Function SendMessage32 Lib "USER32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
 
Declare Function ExtractIcon32 Lib "SHELL32.DLL" Alias _
"ExtractIconA" (ByVal hInst As Long, _
ByVal lpszExeFileName As String, _
ByVal nIconIndex As Long) As Long
 
 
Sub ChangeApplicationIcon()
     
    Dim Icon&
     
     '*****Change Icon To Suit*******
    Const NewIcon$ = "ieshwiz.exe"
     '*****************************
     
    Icon = ExtractIcon32(0, NewIcon, 0)
    SendMessage32 GetActiveWindow32(), &H80, 1, Icon '< 1 = big Icon
    SendMessage32 GetActiveWindow32(), &H80, 0, Icon '< 0 = small Icon
     
End Sub


