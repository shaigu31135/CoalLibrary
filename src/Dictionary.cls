VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'字典类
'作者：马云爱逛京东

'版本：0.1（2023-09-12）

Private Type tDict
    Key As String
    Value As Variant
End Type

Private Dict() As tDict
Private DictSize As Long

'初始化
Private Sub Class_Initialize()
    ReDim Dict(0) As tDict
    DictSize = -1
End Sub

Public Property Get ItemCount() As Long
    ItemCount = DictSize + 1
End Property

'添加项目
Public Sub AddItem(ByVal Key As String, ByVal Value As Variant, Optional ByVal Overwrite As Boolean = False)
    If Key = "" Then Exit Sub
    If FindItem(Key) <> -1 Then
        If Overwrite Then
            ChangeItemByKey Key, Value
        End If
    Else
        DictSize = DictSize + 1
        ReDim Preserve Dict(DictSize) As tDict
        With Dict(DictSize)
            .Key = Key
            .Value = Value
        End With
    End If
End Sub

'移除项目
Public Sub RemoveItemByKey(ByVal Key As String)
    If Key = "" Then Exit Sub
    Dim I As Long, J As Long
    If DictSize = -1 Then Exit Sub
    For I = 0 To UBound(Dict)
        If Dict(I).Key = Key Then
            If DictSize > 0 Then
                DictSize = DictSize - 1
                For J = I To UBound(Dict) - 1
                    Dict(J).Key = Dict(J + 1).Key
                    Dict(J).Value = Dict(J + 1).Value
                Next J
                ReDim Preserve Dict(UBound(Dict) - 1) As tDict
                Exit Sub
            ElseIf DictSize = 0 Then
                DictSize = -1
                Erase Dict
            End If
        End If
    Next I
End Sub

'更改项目（以键方式）
Public Sub ChangeItemByKey(ByVal Key As String, ByVal Value As Variant)
    If Key = "" Then Exit Sub
    If DictSize = -1 Then Exit Sub
    Dim f As Long
    f = FindItem(Key)
    If f = -1 Then Exit Sub
    Dict(f).Value = Value
End Sub

'更改项目（以索引方式）
Public Sub ChangeItemByIndex(ByVal Index As Long, ByVal Value As Variant)
    If DictSize = -1 Then Exit Sub
    If Index > UBound(Dict) Or Index < 0 Then Exit Sub
    Dict(Index).Value = Value
End Sub

'查找项目
'-1表示未找到
Public Function FindItem(ByVal Key As String) As Long
    Dim I As Long, mFlag As Long
    mFlag = -1
    If DictSize = -1 Or Key = "" Then
        FindItem = -1
        Exit Function
    End If
    For I = 0 To UBound(Dict)
        If Dict(I).Key = Key Then
            mFlag = I
            Exit For
        End If
    Next
    FindItem = mFlag
End Function

'获得值的内容（以键方式）
Public Function GetValueByKey(ByVal Key As String) As Variant
    If Key = "" Or DictSize = -1 Then Exit Function
    Dim f As Long
    f = FindItem(Key)
    If f = -1 Then Exit Function
    GetValueByKey = Dict(f).Value
End Function

'获得值的内容（以索引方式）
Public Function GetValueByIndex(ByVal Index As Long) As Variant
    If DictSize = -1 Then Exit Function
    If Index > UBound(Dict) Or Index < 0 Then Exit Function
    GetValueByIndex = Dict(Index).Value
End Function

'清空
Public Sub Clear()
    ReDim Dict(0) As tDict
    DictSize = -1
End Sub

'根据索引获得键名
Public Function GetKey(ByVal Index As Long) As String
    If DictSize = -1 Then Exit Function
    If Index < 0 Or Index > UBound(Dict) Then Exit Function
    GetKey = Dict(Index).Key
End Function

