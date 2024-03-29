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
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
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

'==================================================

Private Sub DrawTimer_Timer()
    '绘制
    EC.Display
    
End Sub

Private Sub Form_Load()
    '初始化Emerald
    StartEmerald Me.Hwnd, 1400, 800
    '创建字体
    MakeFont "微软雅黑"
    '创建页面管理器
    Set EC = New GMan
    
    '创建存档（可选）
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '创建音乐列表
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"
    
    '开始显示
    Me.Show
    DrawTimer.Enabled = True
    
    '在此处初始化你的页面
    Set gamepage1 = New gamepage1
    Set gamepage2 = New gamepage2
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
    
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    If Mouse.state = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
End Sub
