Attribute VB_Name = "Data"
Option Explicit
Public roleX As Long                                                            '��ɫX
Public roleY As Long                                                            ' Y����
Public rolelr As Long                                                           '��ɫ�����߷���
Public HPmax, MPmax As Long                                                     '�������ֵ�������ֵ                                                     '
Public HPnow, MPnow As Long                                                     '��ǰ����ֵ�͵�ǰ����ֵ
Public Level, Expneed, Expnow As Long                                           '�ȼ����������辭�飬��ǰ����
Public Attackstatus As String                                                   '����״̬(attacking,notattacking)
Public OldY As Long, Jumptime As Long                                           '��Y����Ծ��¼ʱ��
Public Copystatus As String                                                     '����״̬���жϽ����״̬(���ǣ�����ѡ����棬����1������2)
Public ChallengeRecord As Long                                                  'ͨ�ع�����¼


Public SEAttack As GMusicList                                                   'SE��Ч������
