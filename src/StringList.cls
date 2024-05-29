VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mList() As String, mIndex As Long, mCount As Long
Private I As Long

Private Sub Class_Initialize()
    mCount = 0
    mIndex = -1
End Sub

'添加项目
Public Sub AddItem(ByVal Item As String)
    If mCount = 0 Then
        ReDim Preserve mList(0)
        mList(0) = Item
        mIndex = 0
        mCount = 1
    ElseIf mCount > 0 Then
        ReDim Preserve mList(mCount)
        mList(mCount) = Item
        mIndex = mCount
        mCount = mCount + 1
    End If
End Sub

'删除项目
Public Sub RemoveItem(ByVal Index As Long)
    If mCount = 0 Or Index < 0 Or Index >= mCount Then Exit Sub
    If Index = mCount - 1 Then                                                  '最后一个项目，直接删除即可
        ReDim Preserve mList(Index - 1) As String
    ElseIf Index < mCount - 1 Then                                              '中间或者开头的项目
        For I = Index To mCount - 2
            mList(I) = mList(I + 1)
        Next I
        mCount = mCount - 1
        ReDim Preserve mList(mCount - 1) As String
        mIndex = -1
    End If
End Sub

'清空项目
Public Sub Clear()
    Erase mList
    mIndex = -1
    mCount = 0
End Sub

'获得项目的数量
Public Function GetItemsCount() As Long
    GetItemsCount = mCount
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

'查找指定名字的项目位置，如果未找到，返回-1；如果列表为空，返回-2，并在Debug输出 <错误：列表为空>
Public Function FindItem(ByVal Find As String) As Long
    Dim Retn As Long
    Retn = -1
    If mCount = 0 Then
        Retn = -2
        Debug.Print "<错误：列表为空>"
    ElseIf mCount > 0 Then
        For mIndex = 0 To mCount - 1
            If Find = mList(mIndex) Then
                Retn = mIndex
                Exit For
            End If
        Next mIndex
    End If
    FindItem = Retn
End Function

'批量添加项目，并以指定的分隔符隔开
Public Sub AddItems(ByVal ItemString As String, Optional ByVal Delimiter As String = "|")
    Dim Tmp As Variant
    Tmp = Split(ItemString, Delimiter)
    For I = 0 To UBound(Tmp)
        AddItem Tmp(I)
    Next I
End Sub

'转换为字符串并输出，如果列表为空，返回空值，并在Debug输出 <错误：列表为空>
Public Function ToString(Optional ByVal Delimiter As String = "|") As String
    Dim Retn As String, mDeli As String * 1
    If Delimiter = "" Then mDeli = "|" Else mDeli = Left(Delimiter, 1)
    If mCount = 0 Then
        Retn = ""
        Debug.Print "<错误：列表为空>"
    ElseIf mCount > 0 Then
        For I = 0 To mCount - 1
            Retn = Retn & mList(I) & mDeli
        Next I
        If Right(Retn, 1) = mDeli Then Retn = Left(Retn, Len(Retn) - 1)
    End If
    ToString = Retn
End Function

'获得指定位置的项目
'如果列表为空，返回空值，并在Debug输出 <错误：列表为空>；如果指定的索引不存在，返回空值，并在Debug输出 <错误：索引不存在>
Public Function List(ByVal Index As Long) As String
    Dim Retn As String
    If mCount = 0 Then
        Retn = ""
        Debug.Print "<错误：列表为空>"
    ElseIf Index < 0 Or Index >= mCount Then
        Retn = ""
        Debug.Print "<错误：索引不存在>"
    Else
        Retn = mList(Index)
    End If
    List = Retn
End Function

'计算所有项目的总和。如果列表为空，返回-2，并在Debug输出 <错误：列表为空>
Public Function Sum() As Double
    Dim Retn As Double
    If mCount = 0 Then
        Retn = -2
        Debug.Print "<错误：列表为空>"
    ElseIf mCount > 0 Then
        For I = 0 To mCount - 1
            Retn = Retn + Val(mList(I))
        Next I
    End If
    Sum = Retn
End Function

'更新项目
'如果列表为空，在Debug输出 <错误：列表为空>；如果指定的索引不存在，在Debug输出 <错误：索引不存在>
Public Sub UpdateItem(ByVal Index As Long, ByVal Item As String)
 If mCount = 0 Then
        Debug.Print "<错误：列表为空>"
    ElseIf Index < 0 Or Index >= mCount Then
        Debug.Print "<错误：索引不存在>"
    Else
        mList(Index) = Item
    End If
End Sub
