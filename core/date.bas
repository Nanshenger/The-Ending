Attribute VB_Name = "Data"
Option Explicit
Public roleX As Long                                                            '角色X
Public roleY As Long                                                            ' Y坐标
Public rolelr As Long                                                           '角色的行走方向
Public HPmax, MPmax As Long                                                     '最大生命值和最大法力值                                                     '
Public HPnow, MPnow As Long                                                     '当前生命值和当前法力值
Public Level, Expneed, Expnow As Long                                           '等级，升级所需经验，当前经验
Public Attackstatus As String                                                   '攻击状态(attacking,notattacking)
Public OldY As Long, Jumptime As Long                                           '旧Y，跳跃记录时间
Public Copystatus As String                                                     '副本状态：判断界面的状态(主城，副本选择界面，副本1，副本2)
Public ChallengeRecord As Long                                                  '通关关数记录


Public SEAttack As GMusicList                                                   'SE音效组声明
