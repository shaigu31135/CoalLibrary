VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'StringBuilder （字符串构造器）

Private mText   As String
Private mLength As Long
Private mReturn As New StringBuilder

'返回字符串
Public Function ToString() As String
    ToString = mText
End Function

'类初始化设置
Private Sub Class_Initialize()
    mText = ""
    mLength = 0
End Sub

'追加字符串
Public Sub AppendText(ByVal Text As String)
    mText = mText & Text
End Sub

'追加字符串到新字符串构造器
Public Function Append(ByVal Text As String) As StringBuilder
    mReturn.Clear
    mReturn.AppendText mText
    mReturn.AppendText Text
    Set Append = mReturn
End Function

'返回长度
Public Property Get Length() As Long
    mLength = Len(mText)
    Length = mLength
End Property

'清空字符串构造器
Public Sub Clear()
    Class_Initialize
End Sub

'转换为大写
Public Sub UpCase()
    mText = UCase(mText)
End Sub

'转换为小写
Public Sub DownCase()
    mText = LCase(mText)
End Sub

'查找是否存在某个字符串
Public Function FindText(ByVal Find As String) As Long
    FindText = VBA.InStr(1, mText, Find, vbTextCompare)
End Function

'从后往前查找字符串
Public Function FindLastText(ByVal Find As String) As Long
    FindLastText = VBA.InStrRev(mText, Find, -1, vbTextCompare)
End Function

'替换字符串
Public Sub ReplaceText(ByVal Find As String, ByVal Replacement As String)
    mText = VBA.Replace(mText, Find, Replacement, 1, -1, vbTextCompare)
End Sub

'替换字符串到新字符串构造器
Public Function Replace(ByVal Find As String, ByVal Replacement As String) As StringBuilder
    mReturn.Clear
    mReturn.AppendText VBA.Replace(mText, Find, Replacement, 1, -1, vbTextCompare)
    Set Replace = mReturn
End Function

'反转字符串
Public Sub ReverseText()
    mText = VBA.StrReverse(mText)
End Sub

'反转字符串到新字符串构造器
Public Function Reverse() As StringBuilder
    mReturn.Clear
    mReturn.AppendText mText
    mReturn.ReverseText
    Set Reverse = mReturn
End Function

'子字符串
Public Function SubString(ByVal Start As Long, Optional Length As Long = -1) As StringBuilder
    mReturn.Clear
    If mText = "" Then
        Set SubString = mReturn
        Exit Function
    Else
        If Length = -1 Then
            mReturn.AppendText VBA.Mid(mText, Start)
        Else
            mReturn.AppendText VBA.Mid(mText, Start, Length)
        End If
        Set SubString = mReturn
    End If
End Function

'覆盖内容
Public Sub RecoverText(ByVal Text As String)
    mText = Text
End Sub
