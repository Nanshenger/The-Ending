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
'   ҳ�������
Dim Page As GPage
'======================================
Option Explicit

Public Sub Update()
    '�����¼�
    Page.Clear
    Page.DrawImage "����.png", 0, 0, Alpha:=1
    Page.DrawImage "���Ǵ��͵�.png", 1300, 760, Alpha:=1, Pos:=posOnCenter
    Page.DrawImage "Ѫ����.png", 20, 20, Alpha:=1
    Page.DrawImage "������.png", 20, 45, Alpha:=1
    Page.DrawImage "������.png", 20, 70, Alpha:=1
    
    If rolelr = 0 And attack = 0 Then Page.DrawImage "��ɫ����1.png", roleX, 670, Alpha:=1, Direction:=DirHorizontal, Pos:=posOnCenter: attacktime = GetTickCount
    If rolelr = 0 And attack = 1 Then Page.DrawImage "��ɫ����2.png", roleX, 670, Alpha:=1, Direction:=DirHorizontal, Pos:=posOnCenter
    If rolelr = 1 And attack = 0 Then Page.DrawImage "��ɫ����1.png", roleX, 670, Alpha:=1, Pos:=posOnCenter: attacktime = GetTickCount
    If rolelr = 1 And attack = 1 Then Page.DrawImage "��ɫ����2.png", roleX, 670, Alpha:=1, Pos:=posOnCenter
    
    
    
    If IsKeyPress(VK_A) Then roleX = roleX - 5: rolelr = 0                      '����A����
    If IsKeyPress(VK_D) Then roleX = roleX + 5: rolelr = 1                      '����D����
    
    If IsKeyPress(VK_J) Then attack = 1                                         ' ����ģ��
    If attack = 0 Then attacktime = GetTickCount
    If attack = 1 And GetTickCount - attacktime >= 500 Then attack = 0
    
    
    If roleX > 1400 Then                                                        '��ֹ�ܳ���Ե
        roleX = 1400
    ElseIf roleX < 0 Then
        roleX = 0
    End If
    If roleX > 1234 And roleX < 1366 Then Page.Writes "����W���븱��", 1155, 520, Color:=argb(255, 255, 200, 0), size:=30
    
    If IsKeyUp(VK_W) And roleX > 1234 And roleX < 1366 Then ECore.NewTransform transFadeIn, 300, "����"
    
    
    
    
    '================================================================
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    Page.Writes "Ѫ��ֵ��", 20, 20, Color:=argb(255, 255, 255, 255)
    Page.Writes "����ֵ��", 20, 45, Color:=argb(255, 255, 255, 255)
    Page.Writes "����ֵ��", 20, 70, Color:=argb(255, 0, 0, 0)
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
'Direction:����Depth�����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(ID As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    '���Page.Res���ǵ�ǰҳ��Я������Դ����������Ϸ��Դ��������ط�����assetsĿ¼�µ�����ͼƬ���뵽��Դ��
    
    'NewImages �½�ͼ����
    '�����������ݲ���ʹ�õĲ�ͬ�в�ͬЧ����
    '·����[ȱʡ]��[ȱʡ]       -����ԭ��С����
    '·�������ű�����[ȱʡ]  -�������ű��ʼ���
    '·�������ȣ��߶�          -����ָ����С����
    Page.Res.NewImages App.path & "\assets"
    '����ҳ��
    ECore.Add Page, "������"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub