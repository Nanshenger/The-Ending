VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
Dim Page As GPage
Dim Crash As New GCrashBox
'======================================
Dim Mapstatus As String                                                         '地图开启状态（open,close）
Dim Challenge As Long                                                           '循环绘制关卡变量
Dim ChallengeTEXT As Long                                                           '循环绘制关卡文字变量
Option Explicit


Public Sub Update()
    '绘制事件
    Page.Clear
    '----------------------------------------------------------------------------------------------------------属性值加载
    If Level = 0 Then Level = 1
    HPnow = HPmax
    MPnow = MPmax
    Copystatus = "主城"
    '----------------------------------------------------------------------------------------------------------主界面图片输出
    Page.DrawImage "主城.jpg", 0, 0
    Page.DrawImage "主城传送点.png", 1220, 720, alpha:=1, pos:=posOnCenter
    '----------------------------------------------------------------------------------------------------------按下W打开副本地图
    '返回最近一个绘图是否和设置为目标的绘图碰撞
    
    Crash.Reset                                                                 '重置，并且设置最近的一个绘图为碰撞目标
    Page.DrawAnimation "Attack", roleX, 630 - roleY, IIf(rolelr = 1, ImgDirection.DirNormal, ImgDirection.DirHorizontal) '<控制人物攻击动画>
    If IsKeyPress(VK_J) And Attackstatus = "notattacking" Then Page.SwitchChannel "Attack", "Default": SEAttack.Play "攻击音效.mp3"
    
    If Crash.IsCrash Then
        Page.Writes "按下W进入副本", 1075, 500, Color:=argb(255, 255, 200, 0), size:=30
        If IsKeyUp(VK_W) And Mapstatus <> "opon " Then Mapstatus = "opon"       '如果在传送阵内且地图状态不为开启，然后地图状态变为开启
    End If
    
    
    If Mapstatus = "opon" Then
        Page.DrawImage "副本地图.png", 50, 130, alpha:=1 '地图绘制
        Page.DrawImage "地图关闭按钮.png", 1225, 220, alpha:=1, pos:=posOnCenter '地图关闭按钮绘制
        If CheckMouse2 = mMouseUp Then Mapstatus = "close"
        For Challenge = 1 To 4 Step 1 '卡关形状绘制
            Page.Paint 0, Challenge * 250 - 100, 180, 200, 100, argb(255, 100, 175, 250)
            If CheckMouse2 = mMouseUp And Challenge <> 1 Then MsgBox "！ "
        Next Challenge
    End If
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    If Mapstatus = "opon" Then
    For ChallengeTEXT = 1 To 4 Step 1 '卡关文字绘制
    Page.Writes "第" & ChallengeTEXT & "关", ChallengeTEXT * 250 - 100, 180, size:=68, Color:=argb(255, 0, 0, 0)
    Next ChallengeTEXT
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
'Direction:方向，Depth：深度
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
Public Sub AnimationMsg(ID As String, msg As String)
    '动画消息接收
    If ID = "Attack" And msg = "attacking" Then Attackstatus = "attacking"
    If ID = "Attack" And msg = "notattacking" Then Attackstatus = "notattacking"
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    '这个Page.Res就是当前页面携带的资源树，管理游戏资源，在这个地方你会把assets目录下的所有图片加入到资源树
    Page.Res.NoLoadDirection = False
    'NewImages 新建图像组
    '参数：（根据参数使用的不同有不同效果）
    '路径，[缺省]，[缺省]       -按照原大小加载
    '路径，缩放比例，[缺省]  -按照缩放比率加载
    '路径，宽度，高度          -按照指定大小加载
    Page.Res.NewImages App.Path & "\assets"
    
    '导入动画
    Page.LoadAnimationsFromDir App.Path & "\animation"
    '创建指定动画
    Page.CreatePlayAnimation "RoleAttack", "Attack", "Reserve"
    
    
    '创建页面
    ECore.Add Page, "主界面"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
