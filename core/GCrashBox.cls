VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCrashBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Emerald ��ش���
'======================================================
'   ��ײ��
'   ������Error 404
'   �汾��1.0.0
'======================================================
    Private Type CrashObj
        tag1 As Integer
        tag2 As Integer
    End Type
    Private Type CrashRet
        Src As CrashObj
        Target As CrashObj
    End Type
    Dim MapW As Long, MapH As Long, Act As Single, oCount As Long
    Dim Map() As Integer, Obj() As CrashObj, Ret() As CrashRet
'======================================================
'   Reset
'   ������ײ��
    Public Sub Reset(w As Long, h As Long, Optional acts As Single = 1)
        Act = acts
        MapW = w * Act: MapH = h * Act
        oCount = 0
        '��ʼ����ײ��ͼ
        ReDim Ret(0), Obj(0), Map(MapW, MapH)
    End Sub
'   Count
'   ȡ����ײ���������
    Public Function Count() As Integer
        Count = UBound(Ret)
    End Function
'   ȡ����ײ���������
    Public Function ObjCount() As Integer
       ObjCount = UBound(Obj)
    End Function
'   ȡ��ָ��������ε�tag1
    Public Function GetObj1(i As Integer) As Integer
        GetObj1 = Obj(i).tag1
    End Function
'   ȡ��ָ��������ε�tag2
    Public Function GetObj2(i As Integer) As Integer
        GetObj2 = Obj(i).tag2
    End Function
'   ȡ��ָ�������ԭʼ���ε�tag1
    Public Function GetSrc1(i As Integer) As Integer
        GetSrc1 = Ret(i).Src.tag1
    End Function
'   ȡ��ָ�������ԭʼ���ε�tag2
    Public Function GetSrc2(i As Integer) As Integer
        GetSrc2 = Ret(i).Src.tag2
    End Function
'   ȡ��ָ���������ײ���ε�tag1
    Public Function GetTarget1(i As Integer) As Integer
        GetTarget1 = Ret(i).Target.tag1
    End Function
'   ȡ��ָ���������ײ���ε�tag2
    Public Function GetTarget2(i As Integer) As Integer
        GetTarget2 = Ret(i).Target.tag2
    End Function
    Public Function CheckCrashRect(mx As Long, my As Long, mw As Long, mh As Long, tag1 As Integer, tag2 As Integer, record As Boolean) As Integer
        Dim Index As Integer

        '��¼����
        ReDim Preserve Obj(UBound(Obj) + 1)
        Index = UBound(Obj)
        With Obj(Index)
            .tag1 = tag1
            .tag2 = tag2
        End With
        
        Dim cx As Long, cy As Long, crashed As Boolean
        Dim rl As Long, rt As Long, rr As Long, rb As Long
        Dim X As Long, Y As Long, h As Long, w As Long
        Dim Step As Integer
        
        '�������ž���
        X = mx * Act: Y = my * Act: w = (mw + 1) * Act: h = (mh + 1) * Act
        
        '��������ı�
        rl = X: rt = Y: rr = X + w: rb = Y + h
        
        '�����α߿���ײ��ͼ
        cy = rt
xTest:
        If cy >= 0 And cy <= MapH Then
            For cx = rl To rr
                If cx >= 0 And cx <= MapW Then
                    If Map(cx, cy) = 0 Then
                        Map(cx, cy) = Index
                    ElseIf Not crashed Then '����õ�ͼ���ѱ���ľ�����䣬˵����ײ
                        GoSub Binggo
                    End If
                End If
            Next
        End If
        Step = Step + 1
        If Step = 1 Then cy = rb: GoTo xTest
        
        cx = rl
yTest:
        If cx >= 0 And cx <= MapW Then
            For cy = rt + 1 To rb - 1
                If cy >= 0 And cy <= MapH Then
                    If Map(cx, cy) = 0 Then
                        Map(cx, cy) = Index
                    ElseIf Not crashed Then
                        GoSub Binggo
                    End If
                End If
            Next
        End If
        Step = Step + 1
        If Step = 3 Then cx = rr: GoTo yTest
        
        Exit Function
        
Binggo:
        If record Then '��¼
            ReDim Preserve Ret(UBound(Ret) + 1)
            With Ret(UBound(Ret))
                .Src = Obj(Index)
                .Target = Obj(Map(cx, cy))
            End With
        End If
        crashed = True: CheckCrashRect = Map(cx, cy): Return
    End Function
    Public Sub AddCrashRect(X As Long, Y As Long, w As Long, h As Long, tag1 As Integer, tag2 As Integer)
        CheckCrashRect X, Y, w, h, tag1, tag2, True
    End Sub