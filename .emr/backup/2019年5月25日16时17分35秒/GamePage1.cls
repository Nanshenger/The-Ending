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
'======================================

Public Sub Update()
    '绘制事件
    Page.Clear
    Page.DrawImage "主城.png", 0, 0, Alpha:=1
    If attack = 0 Then
        Page.DrawImage "角色动作1.png", roleX, 570, Alpha:=1
        timedate = GetTickCount
    End If
    If attack = 1 Then Page.DrawImage "角色动作2.png", roleX, 570, Alpha:=1
    If IsKeyPress(VK_A) Then roleX = roleX - 5                                  '按下A左移
    If IsKeyPress(VK_D) Then roleX = roleX + 5                                  '按下D右移
    If IsKeyPress(VK_J) Then attack = 1
    If attack = 0 Then timedate = GetTickCount
    If attack = 1 And GetTickCount - timedate >= 500 Then attack = 0
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
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
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    '这个Page.Res就是当前页面携带的资源树，管理游戏资源，在这个地方你会把assets目录下的所有图片加入到资源树
    
    'NewImages 新建图像组
    '参数：（根据参数使用的不同有不同效果）
    '路径，[缺省]，[缺省]       -按照原大小加载
    '路径，缩放比例，[缺省]  -按照缩放比率加载
    '路径，宽度，高度          -按照指定大小加载
    Page.Res.NewImages App.Path & "\assets"
    '创建页面
    ECore.Add Page, "主界面"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
