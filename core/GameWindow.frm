VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
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
   StartUpPosition =   2  '��Ļ����
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   ҳ�������
Dim EC As GMan
    
'==================================================
'   �ڴ˴��������ҳ����ģ������
Dim gamepage1 As gamepage1
Dim gamepage2 As gamepage2
Dim TopPage As TopPage
Dim EndMark As Boolean

'==================================================

Private Sub Form_Load()
    '��ʼ��Emerald
    StartEmerald Me.Hwnd, 1300, 750
    '��������
    MakeFont "΢���ź�"
    '����ҳ�������
    Set EC = New GMan
    
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set MusicList = New GMusicList
    MusicList.Create App.Path & "\music"
    
    
    '----------------------------------------------------------------------------------------------------------��Ч
    Set SEAttack = New GMusicList
    SEAttack.Create App.Path & "\music\SE"
    
    
    
    '��ʼ��ʾ
    Me.Show
    'DrawTimer.Enabled = True
    
    '�ڴ˴���ʼ�����ҳ��
    Set gamepage1 = New gamepage1
    Set gamepage2 = New gamepage2
    Set TopPage = New TopPage
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
    '=============================================
    
    '���ûҳ��
    EC.ActivePage = "������"
    '=============================================
    '��ʼ��ֵ
    Attackstatus = "notattacking"
    rolelr = 1
    '----------------------------------------------------------------------------------------------------------����֡���������
    Dim time As Long, FPSTime As Long, FPS As Long, FPSTick As Long, FPSTarget As Long
    Dim LockFPS1 As Long, LockFPS2 As Long, Changed As Boolean
    FPSTime = GetTickCount: time = GetTickCount
    '======================================================================
    '   LockFPS1��������Ŀ��֡�������������
    '   LockFPS2��������Ŀ��֡����֡������ʱ��
    LockFPS1 = 60: LockFPS2 = 50
    '======================================================================����֡�����Ĵ���
    Do While EndMark = False
        '����֡����
        If Changed = False Then
            If FPSctt > 0 And FPS > 0 And GetTickCount - FPSTime > 0 Then
                If FPS > LockFPS2 / 2 Then
                    '�������Դﵽ��FPS����
                    Me.Caption = "����֡����" & 1000 / (FPSctt / FPS)
                    If 1000 / (FPSctt / FPS) < LockFPS1 * 0.8 Then
                        FPSTarget = LockFPS2
                    Else
                        FPSTarget = LockFPS1
                    End If
                End If
                '��̬���ü��
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
    '���������Ϣ
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Mouse.state = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    EndMark = True
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub
