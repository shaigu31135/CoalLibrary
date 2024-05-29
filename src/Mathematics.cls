VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mathematics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Mathematics ����ѧ��

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'Բ����
Public Property Get PI() As Double
    PI = 3.14159265358979
End Property

'��Ȼ������
Public Property Get E() As Double
    E = 2.71828182845905
End Property

'����
Public Function Logarithm(ByVal Base As Double, ByVal Antilogarithm As Double) As Double
    On Error GoTo ExitFunc
    If Base > 0 And Base <> 1 Then
        Logarithm = VBA.Log(Antilogarithm) / VBA.Log(Base)
    Else
        MsgBox "������Base��ֵ����", vbCritical, "������Ϣ"
        Exit Function
    End If
ExitFunc:
    MsgBox "�����������", vbCritical, "������Ϣ"
    Logarithm = 0
End Function

'ָ��
Public Function Exponential(ByVal Base As Double, ByVal Power As Double) As Double
    On Error GoTo ExitFunc
    Exponential = Base ^ Power
    Exit Function
ExitFunc:
    MsgBox "�����������", vbCritical, "������Ϣ"
    Exponential = 0
End Function

'����ת�Ƕ�
Public Function RadToDeg(ByVal Radian As Double) As Double
    RadToDeg = Radian / PI * 180
End Function

'�Ƕ�ת����
Public Function DegToRad(ByVal Degree As Double) As Double
    DegToRad = Degree / 180 * PI
End Function

'���Һ������Ƕȣ�
Public Function SineD(ByVal Degree As Double) As Double
    SineD = SineR(DegToRad(Degree))
End Function

'���Һ������Ƕȣ�
Public Function CosineD(ByVal Degree As Double) As Double
    CosineD = CosineR(DegToRad(Degree))
End Function

'���к������Ƕȣ�
Public Function TangentD(ByVal Degree As Double) As Double
    TangentD = TangentR(DegToRad(Degree))
End Function

'��������Ƕȣ�
Public Function SecantD(ByVal Degree As Double) As Double
    SecantD = SecantR(DegToRad(Degree))
End Function

'�������Ƕȣ�
Public Function CosecantD(ByVal Degree As Double) As Double
    CosecantD = CosecantR(DegToRad(Degree))
End Function

'���к������Ƕȣ�
Public Function CotangentD(ByVal Degree As Double) As Double
    CotangentD = CotangentR(DegToRad(Degree))
End Function

'���Һ��������ȣ�
Public Function SineR(ByVal Radian As Double) As Double
    SineR = VBA.Sin(Radian)
End Function

'���Һ��������ȣ�
Public Function CosineR(ByVal Radian As Double) As Double
    CosineR = VBA.Cos(Radian)
End Function

'���к��������ȣ�
Public Function TangentR(ByVal Radian As Double) As Double
    On Error GoTo ExitFunc
    If CosineR(Radian) = 0 Then
        MsgBox "�����������", vbCritical, "������Ϣ"
        Exit Function
    Else
        TangentR = SineR(Radian) / CosineR(Radian)
    End If
ExitFunc:
    MsgBox "�����������", vbCritical, "������Ϣ"
    TangentR = 0
End Function

'���к��������ȣ�
Public Function CotangentR(ByVal Radian As Double) As Double
    If SineR(Radian) = 0 Then
        MsgBox "�����������", vbCritical, "������Ϣ"
        Exit Function
    Else
        CotangentR = CosineR(Radian) / SineR(Radian)
    End If
End Function

'����������ȣ�
Public Function SecantR(ByVal Radian As Double) As Double
    On Error GoTo ExitFunc
    If CosineR(Radian) = 0 Then
        MsgBox "�����������", vbCritical, "������Ϣ"
        Exit Function
    Else
        SecantR = 1 / CosineR(Radian)
    End If
ExitFunc:
    MsgBox "�����������", vbCritical, "������Ϣ"
    SecantR = 0
End Function

'���������ȣ�
Public Function CosecantR(ByVal Radian As Double) As Double
    On Error GoTo ExitFunc
    If SineR(Radian) = 0 Then
        MsgBox "�����������", vbCritical, "������Ϣ"
        Exit Function
    Else
        CosecantR = 1 / SineR(Radian)
    End If
ExitFunc:
    MsgBox "�����������", vbCritical, "������Ϣ"
    CosecantR = 0
End Function

'˫�����Һ���
Public Function HyperbolicSine(ByVal X As Double) As Double
    HyperbolicSine = (E ^ X - E ^ (-X)) / 2
End Function

'˫�����Һ���
Public Function HyperbolicCosine(ByVal X As Double) As Double
    HyperbolicCosine = (E ^ X + E ^ (-X)) / 2
End Function

'˫�����к���
Public Function HyperbolicTangent(ByVal X As Double) As Double
    HyperbolicTangent = HyperbolicSine(X) / HyperbolicCosine(X)
End Function

'˫�����к���
Public Function HyperbolicCotangent(ByVal X As Double) As Double
    If HyperbolicSine(X) = 0 Then
        MsgBox "�����������", vbCritical, "������Ϣ"
        Exit Function
    Else
        HyperbolicCotangent = HyperbolicCosine(X) / HyperbolicSine(X)
    End If
End Function

'˫�������
Public Function HyperbolicSecant(ByVal X As Double) As Double
    HyperbolicSecant = 1 / HyperbolicCosine(X)
End Function

'˫������
Public Function HyperbolicCosecant(ByVal X As Double) As Double
    If HyperbolicSine(X) = 0 Then
        MsgBox "�����������", vbCritical, "������Ϣ"
        Exit Function
    Else
        HyperbolicCosecant = 1 / HyperbolicSine(X)
    End If
End Function

'�׳�
Public Function Factorial(ByVal X As Integer) As Double
    Dim Retn As Double, I As Integer
    If X < 0 Then
        MsgBox "������X��ֵ����", vbCritical, "������Ϣ"
        Exit Function
    ElseIf X = 0 Then
        Factorial = 1
    ElseIf X >= 1 Then
        Retn = 1
        For I = 1 To X
            Retn = Retn * I
        Next I
        Factorial = Retn
    End If
End Function

'���У��Ӽ����������ȡ����
'CollectionCount - ������
'SampleCount - ������
Public Function Permutation(ByVal CollectionCount As Integer, ByVal SampleCount As Integer) As Double
    If SampleCount > CollectionCount Then
        MsgBox "������SampleCount��ֵ����", vbCritical, "������Ϣ"
        Exit Function
    Else
        Permutation = Factorial(CollectionCount) / Factorial(CollectionCount - SampleCount)
    End If
End Function

'��ϣ��Ӽ����������ȡ������Ȼ����������һ��
'CollectionCount - ������
'SampleCount - ������
Public Function Combination(ByVal CollectionCount As Integer, ByVal SampleCount As Integer) As Double
    If SampleCount > CollectionCount Then
        MsgBox "������SampleCount��ֵ����", vbCritical, "������Ϣ"
        Exit Function
    Else
        Combination = Factorial(CollectionCount) / Factorial(CollectionCount - SampleCount) / Factorial(SampleCount)
    End If
End Function

'���ɽ�����Сֵ�����ֵ֮����������
Public Function RandomInt(ByVal Min As Integer, ByVal Max As Integer) As Integer
    Dim R As Integer
    R = Int(RandomSingle * (Max - Min + 1)) + Min
    RandomInt = R
End Function

'���ɽ�Ϊ����Ľ���[0, 1)�����������
Public Function RandomSingle() As Single
    RandomSingle = Val(Format(Rnd(-timeGetTime), "0.000000000"))
End Function

'���ʼ��
Private Sub Class_Initialize()
    Debug.Print "Mathematics�������Ѿ���ʼ����ɡ�"
End Sub