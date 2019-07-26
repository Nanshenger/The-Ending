VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TOPpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
Dim Page As GPage
'======================================
'----------------------------------------------------------------------------------------------------------音效创建



Public Sub Update()
    
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Dim progress As Single
    
    Page.Clear argb(0, 255, 255, 255)                                           '清空画布
    '-------------------------------------------------------------------------------------------------------人物移动(通用)
    If IsKeyPress(VK_A) And Attackstatus = "notattacking" Then roleX = roleX - 5: rolelr = 0 '按下A左移
    If IsKeyPress(VK_D) And Attackstatus = "notattacking" Then roleX = roleX + 5: rolelr = 1 '按下D右移
    If roleX > 1400 Then roleX = 1400                                           '防止跑出右侧边缘
    If roleX < 0 Then roleX = 0                                                 '防止跑出左侧边缘
    '-------------------------------------------------------------------------------------------------------人物属性值计算(通用)
    HPmax = 100 + 10 * Level                                                    '血量计算
    MPmax = 50 + 5 * Level                                                      '法力值计算
    Expneed = Int(100 * 1.1 ^ Level)                                            '升级所需总经验值计算
    
    '-------------------------------------------------------------------------------------------------------人物属性输出(通用)
    Page.DrawImage "血量条.png", 145, 20, cw:=HPnow / HPmax * 145 + 5, alpha:=1
    Page.DrawImage "蓝量条.png", 145, 45, cw:=MPnow / MPmax * 145 + 5, alpha:=1
    Page.DrawImage "经验条.png", 145, 70, cw:=Expnow / Expneed * 145 + 5, alpha:=1
    Page.DrawImage "头像边框.png", 5, 5, alpha:=1
    Page.DrawImage "头像.png", 34, 26, alpha:=1
    '-------------------------------------------------------------------------------------------------------人物跳跃代码(通用)
    '如果jumptime为0（默认为0）而且K按下，则向jumptime赋值
    If IsKeyPress(VK_K) And Jumptime = 0 Then Jumptime = GetTickCount
    '如果jumptime被赋值
    If Jumptime <> 0 Then
        'jumptime进度
        progress = (GetTickCount - Jumptime) / 300                              '300=跳下和跳下所用时间
        If progress > 1 Then progress = 2 - progress                            '如果进度>1，说明开始向下
        If progress < 0 Then progress = 0: Jumptime = 0                         '如果进度<0，说明已经向下完成
        roleY = OldY + (200 * (progress ^ 2))                                   ' 100 = jump height 设置高度
    End If
    '-------------------------------------------------------------------------------------------------------游戏功能图标输出
    Page.DrawImage "功能背包.png", 20, 680, alpha:=1
    Page.DrawImage "功能技能.png", 85, 680, alpha:=1
    Page.DrawImage "功能元神.png", 150, 680, alpha:=1
    Page.DrawImage "功能任务.png", 215, 680, alpha:=1
    
    
    
    
    
    '================================================================
    '   示例：绘制文字I
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================人物属性输出（文字，通用）
    Page.Writes "血量值：" & HPnow & "/" & HPmax, 145, 20, Color:=argb(255, 255, 255, 255) '当前/最大
    Page.Writes "法力值：" & MPnow & "/" & MPmax, 145, 45, Color:=argb(255, 255, 255, 255) '当前/最大
    Page.Writes "经验值：" & Expnow & "/" & Expneed, 145, 70, Color:=argb(255, 0, 0, 0) '当前/最大
    Page.Writes "等级：" & Level, 34, 100, Color:=argb(255, 255, 0, 255), size:=20 '等级显示
    Page.Writes Str(HPnow / HPmax * 100) & "%", 300, 20, Color:=argb(255, 0, 0, 0) '血量百分比
    Page.Writes Str(MPnow / MPmax * 100) & "%", 300, 45, Color:=argb(255, 0, 0, 0) '法力百分比
    Page.Writes Str(Expnow / Expneed * 100) & "%", 300, 70, Color:=argb(255, 0, 0, 0) '经验百分比
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
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
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.TopPage = True
    Page.Res.NewImages App.Path & "\assets"
    '创建页面
    ECore.Add Page, "[你的页面名称]"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
