VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "窗口名称"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
Dim EC As GMan
    
'==================================================
'   在此处放置你的页面类模块声明
Dim gamepage1 As gamepage1
Dim gamepage2 As gamepage2
Dim TopPage As TopPage
Dim EndMark As Boolean

'==================================================

Private Sub Form_Load()
    '初始化Emerald
    StartEmerald Me.Hwnd, 1300, 750
    '创建字体
    MakeFont "微软雅黑"
    '创建页面管理器
    Set EC = New GMan
    
    '创建存档（可选）
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '创建音乐列表
    Set MusicList = New GMusicList
    MusicList.Create App.Path & "\music"
    
    
    '----------------------------------------------------------------------------------------------------------音效
    Set SEAttack = New GMusicList
    SEAttack.Create App.Path & "\music\SE"
    
    
    
    '开始显示
    Me.Show
    'DrawTimer.Enabled = True
    
    '在此处初始化你的页面
    Set gamepage1 = New gamepage1
    Set gamepage2 = New gamepage2
    Set TopPage = New TopPage
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
    '=============================================
    
    '设置活动页面
    EC.ActivePage = "主界面"
    '=============================================
    '初始数值
    Attackstatus = "notattacking"
    rolelr = 1
    '----------------------------------------------------------------------------------------------------------锁定帧数定义变量
    Dim time As Long, FPSTime As Long, FPS As Long, FPSTick As Long, FPSTarget As Long
    Dim LockFPS1 As Long, LockFPS2 As Long, Changed As Boolean
    FPSTime = GetTickCount: time = GetTickCount
    '======================================================================
    '   LockFPS1：锁定的目标帧数（正常情况）
    '   LockFPS2：锁定的目标帧数（帧数不够时）
    LockFPS1 = 60: LockFPS2 = 50
    '======================================================================锁定帧数核心代码
    Do While EndMark = False
        '锁定帧数！
        If Changed = False Then
            If FPSctt > 0 And FPS > 0 And GetTickCount - FPSTime > 0 Then
                If FPS > LockFPS2 / 2 Then
                    '分析可以达到的FPS档次
                    Me.Caption = "分析帧数：" & 1000 / (FPSctt / FPS)
                    If 1000 / (FPSctt / FPS) < LockFPS1 * 0.8 Then
                        FPSTarget = LockFPS2
                    Else
                        FPSTarget = LockFPS1
                    End If
                End If
                '动态设置间隔
                If FPSTarget > 0 Then FPSTick = (1000 / FPSTarget) / ((((GetTickCount - FPSTime) / FPS) * FPSTarget) / 100): Changed = True
            End If
            If FPSTick < 0 Then FPSTick = 0
        End If
        If GetTickCount - FPSTime >= 1000 Then
            FPSTime = GetTickCount
            FPS = 0
        End If
        If GetTickCount - time >= FPSTick Then
            time = GetTickCount: FPS = FPS + 1: Changed = False
            EC.Display
        End If
        DoEvents
    Loop
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Mouse.state = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    EndMark = True
    '释放Emerald资源
    EndEmerald
End Sub
