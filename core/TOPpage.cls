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
'   ҳ�������
Dim Page As GPage
'======================================
'----------------------------------------------------------------------------------------------------------��Ч����



Public Sub Update()
    
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Dim progress As Single
    
    Page.Clear argb(0, 255, 255, 255)                                           '��ջ���
    '-------------------------------------------------------------------------------------------------------�����ƶ�(ͨ��)
    If IsKeyPress(VK_A) And Attackstatus = "notattacking" Then roleX = roleX - 5: rolelr = 0 '����A����
    If IsKeyPress(VK_D) And Attackstatus = "notattacking" Then roleX = roleX + 5: rolelr = 1 '����D����
    If roleX > 1400 Then roleX = 1400                                           '��ֹ�ܳ��Ҳ��Ե
    If roleX < 0 Then roleX = 0                                                 '��ֹ�ܳ�����Ե
    '-------------------------------------------------------------------------------------------------------��������ֵ����(ͨ��)
    HPmax = 100 + 10 * Level                                                    'Ѫ������
    MPmax = 50 + 5 * Level                                                      '����ֵ����
    Expneed = Int(100 * 1.1 ^ Level)                                            '���������ܾ���ֵ����
    
    '-------------------------------------------------------------------------------------------------------�����������(ͨ��)
    Page.DrawImage "Ѫ����.png", 145, 20, cw:=HPnow / HPmax * 145 + 5, alpha:=1
    Page.DrawImage "������.png", 145, 45, cw:=MPnow / MPmax * 145 + 5, alpha:=1
    Page.DrawImage "������.png", 145, 70, cw:=Expnow / Expneed * 145 + 5, alpha:=1
    Page.DrawImage "ͷ��߿�.png", 5, 5, alpha:=1
    Page.DrawImage "ͷ��.png", 34, 26, alpha:=1
    '-------------------------------------------------------------------------------------------------------������Ծ����(ͨ��)
    '���jumptimeΪ0��Ĭ��Ϊ0������K���£�����jumptime��ֵ
    If IsKeyPress(VK_K) And Jumptime = 0 Then Jumptime = GetTickCount
    '���jumptime����ֵ
    If Jumptime <> 0 Then
        'jumptime����
        progress = (GetTickCount - Jumptime) / 300                              '300=���º���������ʱ��
        If progress > 1 Then progress = 2 - progress                            '�������>1��˵����ʼ����
        If progress < 0 Then progress = 0: Jumptime = 0                         '�������<0��˵���Ѿ��������
        roleY = OldY + (200 * (progress ^ 2))                                   ' 100 = jump height ���ø߶�
    End If
    '-------------------------------------------------------------------------------------------------------��Ϸ����ͼ�����
    Page.DrawImage "���ܱ���.png", 20, 680, alpha:=1
    Page.DrawImage "���ܼ���.png", 85, 680, alpha:=1
    Page.DrawImage "����Ԫ��.png", 150, 680, alpha:=1
    Page.DrawImage "��������.png", 215, 680, alpha:=1
    
    
    
    
    
    '================================================================
    '   ʾ������������I
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================����������������֣�ͨ�ã�
    Page.Writes "Ѫ��ֵ��" & HPnow & "/" & HPmax, 145, 20, Color:=argb(255, 255, 255, 255) '��ǰ/���
    Page.Writes "����ֵ��" & MPnow & "/" & MPmax, 145, 45, Color:=argb(255, 255, 255, 255) '��ǰ/���
    Page.Writes "����ֵ��" & Expnow & "/" & Expneed, 145, 70, Color:=argb(255, 0, 0, 0) '��ǰ/���
    Page.Writes "�ȼ���" & Level, 34, 100, Color:=argb(255, 255, 0, 255), size:=20 '�ȼ���ʾ
    Page.Writes Str(HPnow / HPmax * 100) & "%", 300, 20, Color:=argb(255, 0, 0, 0) 'Ѫ���ٷֱ�
    Page.Writes Str(MPnow / MPmax * 100) & "%", 300, 45, Color:=argb(255, 0, 0, 0) '�����ٷֱ�
    Page.Writes Str(Expnow / Expneed * 100) & "%", 300, 70, Color:=argb(255, 0, 0, 0) '����ٷֱ�
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '������Ϣ����

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
    Page.TopPage = True
    Page.Res.NewImages App.Path & "\assets"
    '����ҳ��
    ECore.Add Page, "[���ҳ������]"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub