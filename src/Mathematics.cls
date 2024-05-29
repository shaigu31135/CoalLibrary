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

'Mathematics （数学）

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'圆周率
Public Property Get PI() As Double
    PI = 3.14159265358979
End Property

'自然对数底
Public Property Get E() As Double
    E = 2.71828182845905
End Property

'对数
Public Function Logarithm(ByVal Base As Double, ByVal Antilogarithm As Double) As Double
    On Error GoTo ExitFunc
    If Base > 0 And Base <> 1 Then
        Logarithm = VBA.Log(Antilogarithm) / VBA.Log(Base)
    Else
        MsgBox "参数：Base的值错误。", vbCritical, "错误信息"
        Exit Function
    End If
ExitFunc:
    MsgBox "计算结果溢出。", vbCritical, "错误信息"
    Logarithm = 0
End Function

'指数
Public Function Exponential(ByVal Base As Double, ByVal Power As Double) As Double
    On Error GoTo ExitFunc
    Exponential = Base ^ Power
    Exit Function
ExitFunc:
    MsgBox "计算结果溢出。", vbCritical, "错误信息"
    Exponential = 0
End Function

'弧度转角度
Public Function RadToDeg(ByVal Radian As Double) As Double
    RadToDeg = Radian / PI * 180
End Function

'角度转弧度
Public Function DegToRad(ByVal Degree As Double) As Double
    DegToRad = Degree / 180 * PI
End Function

'正弦函数（角度）
Public Function SineD(ByVal Degree As Double) As Double
    SineD = SineR(DegToRad(Degree))
End Function

'余弦函数（角度）
Public Function CosineD(ByVal Degree As Double) As Double
    CosineD = CosineR(DegToRad(Degree))
End Function

'正切函数（角度）
Public Function TangentD(ByVal Degree As Double) As Double
    TangentD = TangentR(DegToRad(Degree))
End Function

'正割函数（角度）
Public Function SecantD(ByVal Degree As Double) As Double
    SecantD = SecantR(DegToRad(Degree))
End Function

'余割函数（角度）
Public Function CosecantD(ByVal Degree As Double) As Double
    CosecantD = CosecantR(DegToRad(Degree))
End Function

'余切函数（角度）
Public Function CotangentD(ByVal Degree As Double) As Double
    CotangentD = CotangentR(DegToRad(Degree))
End Function

'正弦函数（弧度）
Public Function SineR(ByVal Radian As Double) As Double
    SineR = VBA.Sin(Radian)
End Function

'余弦函数（弧度）
Public Function CosineR(ByVal Radian As Double) As Double
    CosineR = VBA.Cos(Radian)
End Function

'正切函数（弧度）
Public Function TangentR(ByVal Radian As Double) As Double
    On Error GoTo ExitFunc
    If CosineR(Radian) = 0 Then
        MsgBox "计算结果溢出。", vbCritical, "错误信息"
        Exit Function
    Else
        TangentR = SineR(Radian) / CosineR(Radian)
    End If
ExitFunc:
    MsgBox "计算结果溢出。", vbCritical, "错误信息"
    TangentR = 0
End Function

'余切函数（弧度）
Public Function CotangentR(ByVal Radian As Double) As Double
    If SineR(Radian) = 0 Then
        MsgBox "计算结果溢出。", vbCritical, "错误信息"
        Exit Function
    Else
        CotangentR = CosineR(Radian) / SineR(Radian)
    End If
End Function

'正割函数（弧度）
Public Function SecantR(ByVal Radian As Double) As Double
    On Error GoTo ExitFunc
    If CosineR(Radian) = 0 Then
        MsgBox "计算结果溢出。", vbCritical, "错误信息"
        Exit Function
    Else
        SecantR = 1 / CosineR(Radian)
    End If
ExitFunc:
    MsgBox "计算结果溢出。", vbCritical, "错误信息"
    SecantR = 0
End Function

'余割函数（弧度）
Public Function CosecantR(ByVal Radian As Double) As Double
    On Error GoTo ExitFunc
    If SineR(Radian) = 0 Then
        MsgBox "计算结果溢出。", vbCritical, "错误信息"
        Exit Function
    Else
        CosecantR = 1 / SineR(Radian)
    End If
ExitFunc:
    MsgBox "计算结果溢出。", vbCritical, "错误信息"
    CosecantR = 0
End Function

'双曲正弦函数
Public Function HyperbolicSine(ByVal X As Double) As Double
    HyperbolicSine = (E ^ X - E ^ (-X)) / 2
End Function

'双曲余弦函数
Public Function HyperbolicCosine(ByVal X As Double) As Double
    HyperbolicCosine = (E ^ X + E ^ (-X)) / 2
End Function

'双曲正切函数
Public Function HyperbolicTangent(ByVal X As Double) As Double
    HyperbolicTangent = HyperbolicSine(X) / HyperbolicCosine(X)
End Function

'双曲余切函数
Public Function HyperbolicCotangent(ByVal X As Double) As Double
    If HyperbolicSine(X) = 0 Then
        MsgBox "计算结果溢出。", vbCritical, "错误信息"
        Exit Function
    Else
        HyperbolicCotangent = HyperbolicCosine(X) / HyperbolicSine(X)
    End If
End Function

'双曲正割函数
Public Function HyperbolicSecant(ByVal X As Double) As Double
    HyperbolicSecant = 1 / HyperbolicCosine(X)
End Function

'双曲余割函数
Public Function HyperbolicCosecant(ByVal X As Double) As Double
    If HyperbolicSine(X) = 0 Then
        MsgBox "计算结果溢出。", vbCritical, "错误信息"
        Exit Function
    Else
        HyperbolicCosecant = 1 / HyperbolicSine(X)
    End If
End Function

'阶乘
Public Function Factorial(ByVal X As Integer) As Double
    Dim Retn As Double, I As Integer
    If X < 0 Then
        MsgBox "参数：X的值错误。", vbCritical, "错误信息"
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

'排列：从集合中逐个抽取样本
'CollectionCount - 集合数
'SampleCount - 样本数
Public Function Permutation(ByVal CollectionCount As Integer, ByVal SampleCount As Integer) As Double
    If SampleCount > CollectionCount Then
        MsgBox "参数：SampleCount的值错误。", vbCritical, "错误信息"
        Exit Function
    Else
        Permutation = Factorial(CollectionCount) / Factorial(CollectionCount - SampleCount)
    End If
End Function

'组合：从集合中逐个抽取样本，然后将样本视作一组
'CollectionCount - 集合数
'SampleCount - 样本数
Public Function Combination(ByVal CollectionCount As Integer, ByVal SampleCount As Integer) As Double
    If SampleCount > CollectionCount Then
        MsgBox "参数：SampleCount的值错误。", vbCritical, "错误信息"
        Exit Function
    Else
        Combination = Factorial(CollectionCount) / Factorial(CollectionCount - SampleCount) / Factorial(SampleCount)
    End If
End Function

'生成介于最小值和最大值之间的随机整数
Public Function RandomInt(ByVal Min As Integer, ByVal Max As Integer) As Integer
    Dim R As Integer
    R = Int(RandomSingle * (Max - Min + 1)) + Min
    RandomInt = R
End Function

'生成较为无序的介于[0, 1)的随机浮点数
Public Function RandomSingle() As Single
    RandomSingle = Val(Format(Rnd(-timeGetTime), "0.000000000"))
End Function

'类初始化
Private Sub Class_Initialize()
    Debug.Print "Mathematics工具类已经初始化完成。"
End Sub
