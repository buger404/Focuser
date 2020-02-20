Attribute VB_Name = "Hidding"
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_VISIBLE = &H10000000
Public Const GWL_STYLE = (-16)
Public Const SW_MINIMIZE = 6
Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim l As Long
    Dim ld As Long
    ld = GetWindowLong(hwnd, GWL_STYLE)
    If ((ld And WS_MINIMIZEBOX) = WS_MINIMIZEBOX) Then
        If ((ld And WS_VISIBLE)) Then
            If ShowWindow(hwnd, SW_MINIMIZE) Then
            End If
        End If
    End If
    EnumWindowsProc = True
End Function

