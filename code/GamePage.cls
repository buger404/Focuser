VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim EndHour As Long, EndMinute As Long, EndPercent As Long
    Dim OtherTick As Long, FocusTick As Long, StartTime As Long
    Dim EndHwnd As Long, ForceMode As Long, EndString As String * 255
    Dim MY As Long, MData As Long
    Dim Step As Long, lP As POINTAPI
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(1, 0, 0, 0) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Select Case Step
        Case 0: Call Step0
        Case 1: Call Step1
        Case 2: Call Step2
        Case 3: Call Step3
        Case 4: Call Step4
    End Select
    
    If Step <> 3 Then
        If IsKeyPress(VK_ESCAPE) Then
            Unload GameWindow
            End
        End If
    End If
    
    If Step = 3 And ForceMode = 0 Then
        If IsKeyPress(VK_ESCAPE) Then
            GameWindow.Move Screen.Width / 2 - GameWindow.Width / 2, Screen.Height / 2 - GameWindow.Height / 2
            StartTime = GetTickCount: Step = 0: ECore.NewTransform transDarkTo
        End If
    End If
End Sub
Sub Step0()
    Dim m As MButtonState
    
    Page.Paint 0, 0, 0, GW + 1, 50, argb(40, 0, 0, 0)
    Page.Writes "专注结束时间", GW / 2, 10, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    
    Page.Writes format(EndHour, "00"), 0, 60, 24, IIf(MY <> -1 And Mouse.X <= GW / 2, argb(255, 254, 205, 1), argb(200, 255, 255, 255)), GW / 2 - 10, 35, align:=StringAlignmentCenter, style:=FontStyleBold
    If MY <> -1 Then
        m = CheckMouse(0, 0, GW / 2, GH)
    Else
        m = CheckMouse2
    End If
    
    If m = mMouseDown Then
        If MY = -1 Then MY = Mouse.y: MData = EndHour
        EndHour = ((MData - (((Mouse.y - MY) / 100 * 24) Mod 24) + 24) Mod 24)
    ElseIf m = mMouseUp Then
        If Mouse.y = MY Or MY = -1 Then EndHour = EndHour + 1
        MY = -1
    End If

    If EndHour >= 24 Then EndHour = 0
    
    Page.Writes ":", GW / 2 - 10, 58, 24, argb(255, 255, 255, 255), 20, 35, align:=StringAlignmentCenter, style:=FontStyleBold
    Page.Writes format(EndMinute, "00"), GW / 2 + 10, 60, 24, IIf(MY <> -1 And Mouse.X > GW / 2, argb(255, 254, 205, 1), argb(200, 255, 255, 255)), GW / 2 - 10, 35, align:=StringAlignmentCenter, style:=FontStyleBold
    If MY <> -1 Then
        m = CheckMouse(GW / 2, 0, GW / 2, GH)
    Else
        m = CheckMouse2
    End If
    
    If m = mMouseDown Then
        If MY = -1 Then MY = Mouse.y: MData = EndMinute
        EndMinute = ((MData - (((Mouse.y - MY) / 100 * 60) Mod 60) + 60) Mod 60)
    ElseIf m = mMouseUp Then
        If Mouse.y = MY Or MY = -1 Then EndMinute = EndMinute + 5
        MY = -1
    End If
    
    If EndMinute >= 60 Then EndMinute = 0
    
    Page.Paint 0, 0, GH - 39, GW + 1, 40, argb(255, 24, 180, 255)
    If CheckMouse2 = mMouseUp Then Step = 1: ECore.NewTransform transDarkTo
    Page.Writes "设置", GW / 2, GH - 35, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
End Sub
Sub Step1()
    Page.Paint 0, 0, 0, GW + 1, 50, argb(40, 0, 0, 0)
    Page.Writes "专注百分比", GW / 2, 10, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    
    Page.Writes EndPercent & "%", 0, 60, 24, IIf(MY <> -1, argb(255, 254, 205, 1), argb(200, 255, 255, 255)), GW, 35, align:=StringAlignmentCenter, style:=FontStyleBold
    If MY <> -1 Then
        m = CheckMouse(0, 0, GW, GH)
    Else
        m = CheckMouse2
    End If
    
    If m = mMouseDown Then
        If MY = -1 Then MY = Mouse.y: MData = EndPercent
        EndPercent = ((MData - (((Mouse.y - MY) / 100 * 100) Mod 100) + 100) Mod 100)
    ElseIf m = mMouseUp Then
        If Mouse.y = MY Or MY = -1 Then EndPercent = EndPercent + 5
        MY = -1
    End If
    
    If EndPercent > 100 Then EndPercent = 0
    
    Page.Paint 0, 0, GH - 39, GW + 1, 40, argb(255, 24, 180, 255)
    If CheckMouse2 = mMouseUp Then StartTime = GetTickCount: Step = 4: ECore.NewTransform transDarkTo
    Page.Writes "设置", GW / 2, GH - 35, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
End Sub
Sub Step2()
    Page.Paint 0, 0, 0, GW + 1, 50, argb(40, 0, 0, 0)
    Page.Writes "专注窗口", GW / 2, 10, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    Dim p As POINTAPI, Title As String * 255
    On Error Resume Next
    GetCursorPos p
    GetWindowTextA GetParenter(WindowFromPoint(p.X, p.y)), Title, 255
    
    If 6 - Int((GetTickCount - StartTime) / 1000) <= 0 Then
        Dim Hwnd As Long
        GetCursorPos p
        Hwnd = WindowFromPoint(p.X, p.y)
        Do While GetParent(Hwnd) <> 0
            Hwnd = GetParent(Hwnd)
            DoEvents
        Loop
        EndString = Title
        EndHwnd = Hwnd: OtherTick = 0: FocusTick = 0
        GameWindow.Move Screen.Width - 10 * Screen.TwipsPerPixelX - GameWindow.Width, Screen.Height - 10 * Screen.TwipsPerPixelY - GameWindow.Height - GetTaskbarHeight
        StartTime = GetTickCount: Step = 3: ECore.NewTransform transDarkTo
        Exit Sub
    End If
    
    If WindowFromPoint(p.X, p.y) = GameWindow.Hwnd Then
        StartTime = GetTickCount
        Page.Writes "当前窗口无效", 0, 50, 14, argb(255, 255, 255, 255), GW, 93, align:=StringAlignmentCenter, style:=FontStyleBold
    Else
        Page.Writes 6 - Int((GetTickCount - StartTime) / 1000) & "s后设置鼠标所在的窗口为专注窗口" & vbCrLf & "当前：" & vbCrLf & Replace(Title, Chr(0), ""), 0, 50, 14, argb(255, 255, 255, 255), GW, 93, align:=StringAlignmentCenter, style:=FontStyleBold
    End If
End Sub
Sub Step4()
    Page.Paint 0, 0, 0, GW + 1, 50, argb(40, 0, 0, 0)
    Page.Writes "强迫模式", GW / 2, 10, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    Page.Writes IIf(ForceMode = 1, "ON", "OFF"), 0, 60, 24, argb(255, 255, 255, 255), GW, 35, align:=StringAlignmentCenter, style:=FontStyleBold
    If CheckMouse2 = mMouseUp Then ForceMode = IIf(ForceMode = 1, 0, 1)
    
    Page.Paint 0, 0, GH - 39, GW + 1, 40, argb(255, 24, 180, 255)
    If CheckMouse2 = mMouseUp Then StartTime = GetTickCount: Step = 2: ECore.NewTransform transDarkTo
    Page.Writes "设置", GW / 2, GH - 35, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
End Sub
Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelY) - rectVal.Bottom) * Screen.TwipsPerPixelY
End Function
Sub Step3()
    Dim Per As Long, Title As String * 255, Hwnd As Long
    On Error Resume Next
    Hwnd = GetParenter(GetForegroundWindow)
    GetWindowTextA Hwnd, Title, 255
    
    If Hwnd = EndHwnd And Title = EndString Then
        Dim p As POINTAPI
        GetCursorPos p
        If GetParenter(WindowFromPoint(p.X, p.y)) <> EndHwnd Then
            OtherTick = OtherTick + 1
        Else
            lP = p
            FocusTick = FocusTick + 1
        End If
    Else
        OtherTick = OtherTick + 1
    End If
    Per = Int(FocusTick / (FocusTick + OtherTick) * 100)
    
    Dim Ta As Long, No As Long
    Ta = Val(format(EndHour, "00") & format(EndMinute, "00"))
    No = Val(format(Hour(Now), "00") & format(Minute(Now), "00"))
    
    If No >= Ta Then
        GameWindow.Move Screen.Width / 2 - GameWindow.Width / 2, Screen.Height / 2 - GameWindow.Height / 2
        StartTime = GetTickCount: Step = 0: ECore.NewTransform transDarkTo
    End If
   
    If Per < EndPercent Then
        Page.Paint 0, 0, GH - (GH * Per / 100), GW + 1, GH * Per / 100 + 1, argb(100, 253, 86, 60)
        If ForceMode Then
            Dim l As Long
            l = EnumWindows(AddressOf EnumWindowsProc, 0)
            ShowWindow EndHwnd, SW_SHOW
            SetCursorPos lP.X, lP.y
            SetForegroundWindow EndHwnd
        End If
    ElseIf Per < EndPercent + (100 - EndPercent) / 2 Then
        Page.Paint 0, 0, GH - (GH * Per / 100), GW + 1, GH * Per / 100 + 1, argb(100, 255, 170, 49)
    Else
        Page.Paint 0, 0, GH - (GH * Per / 100), GW + 1, GH * Per / 100 + 1, argb(100, 9, 241, 117)
    End If
    Page.Writes Per & "%", 0, 60, 24, argb(255, 255, 255, 255), GW, 35, align:=StringAlignmentCenter, style:=FontStyleBold

    'Sleep 100
End Sub
Function GetParenter(Hwnd As Long) As Long
    Dim h As Long
    h = Hwnd
    Do While GetParent(h) > 0
        h = GetParent(h)
    Loop
    GetParenter = h
End Function
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets"
    '创建页面
    ECore.Add Page, "GamePage"
    EndPercent = 85
    MY = -1
    EndHour = Hour(Now): EndMinute = Int(Minute(Now) / 5) * 5
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
