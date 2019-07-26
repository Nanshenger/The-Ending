VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
Dim Page As GPage
'======================================
Option Explicit

Public Sub Update()
    '绘制事件
    Page.Clear
    '----------------------------------------------------------------------------------------------------------属性值加载
    
    Copystatus = "副本选择"
    '----------------------------------------------------------------------------------------------------------副本选择背景绘制
    If Copystatus = "副本选择" Then Page.DrawImage "副本选择界面背景.jpg", 0, 0, alpha:=1
    
    '----------------------------------------------------------------------------------------------------------控制人物攻击动画
    Page.DrawAnimation "Attack", roleX, 630 - roleY, IIf(rolelr = 1, ImgDirection.DirNormal, ImgDirection.DirHorizontal)
    If IsKeyPress(VK_J) And Attackstatus = "notattacking" Then Page.SwitchChannel "Attack", "Default": SEAttack.Play "攻击音效.mp3"
    '切换动画的通道，切换的时候，动画会重新播放，所以就利用切换到目前通道，来实现重新播放
    
    
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
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.Path & "\assets"
    '创建页面
    ECore.Add Page, "副本"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    '导入动画
    Page.LoadAnimationsFromDir App.Path & "\animation"
    '创建指定动画
    Page.CreatePlayAnimation "RoleAttack", "Attack", "Reserve"
    '===============================================
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '动画消息接收
    If ID = "Attack" And msg = "attacking" Then Attackstatus = "attacking"
    If ID = "Attack" And msg = "notattacking" Then Attackstatus = "notattacking"
End Sub
