Attribute VB_Name = "date"
Option Explicit
Public roleX As Long
Public roleY As Long

Public rolelr As Long                                                           '角色的方向
Public attack As Long                                                           '攻击状态
Public attacktime As Long                                                       '时间记录
Public HPmax, MPmax As Long                                                     '最大生命值和最大法力值                                                     '
Public HPnow, MPnow As Long                                                     '当前生命值和当前法力值
Public level, needexp, nowexp As Long                                           '等级，升级所需经验，当前经验

Public Attackstatus As String
