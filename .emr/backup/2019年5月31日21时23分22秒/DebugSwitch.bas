Attribute VB_Name = "DebugSwitch"
'   Emerald ������

'======================================================
'   �Ƿ���Debug
Public Const DebugMode As Boolean = True
'   ���ÿ���LOGO
Public Const DisableLOGO As Boolean = False
'   �Ƿ���������Ŀ���LOGO�������Դ�Ѿ�������ϣ�
Public Const HideLOGO As Boolean = False
'   �����¼��ʱ�����죩
Public Const UpdateCheckInterval As Long = 1
'   ���¼�鳬ʱʱ�䣨���룩
Public Const UpdateTimeOut As Long = 2000
'======================================================


'==============================================================================
'   �汾����ע������
'==============================================================================
'   1.������޸�
'      ���������Ϸ�����ҵ����´���
'       Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'           If Mouse.State = 0 Then UpdateMouse X, Y, 0, button
'       End Sub
'      *****�����޸�Ϊ��
'       Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'           If Mouse.State = 0 Then
'               UpdateMouse X, Y, 0, button
'           Else
'               Mouse.X = X: Mouse.Y = Y
'           End If
'       End Sub
'   2.������ջ����޸�
'     ������Ļ�ͼ���̼��룺
'       Page.Clear
'==============================================================================
'   1.��Դ���صĸı�
'     ���Page.NewImagesǨ�Ƶ�Page.Res.NewImages
'==============================================================================
'   1.���ش���ĸı�
'     ���ڿ���LOGO�ļ��룬
'     �����������ҳ��Ϳ�������Timer�Ĵ���ת�Ƶ�����ҳ��֮ǰ��������һ�С�Me.Show��
'   *�ò��������Բ���Emerald�ṩ��ģ��
'   2.Emerald��ʼ���ĸı�
'     �����뵽Emerald�Ĵ��ڴ�С��������Emerald��������һ�Ρ�
'==============================================================================




'======================================================
'   Warning: please don't edit the following code .
Public Debug_focus As Boolean
'======================================================
