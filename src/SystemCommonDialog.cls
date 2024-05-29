VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'通用对话框
'此类模块依赖comdlg32.dll组件
'改编自作者：Donald Grover的源代码。
'修改：马云爱逛京东

'版本：0.1

'常量
Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_ROMAN = 16
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const REGULAR_FONTTYPE = &H400
Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const PD_PRINTSETUP = &H40
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const MAXFILECOUNT = 32767&
Private Const LOGPIXELSY = 90
'结构体

'点
Private Type PointAPI
    X As Long
    Y As Long
End Type

'矩形
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'打开文件对话框
Private Type OpenFileName
    lStructSize As Long                                                         '结构体的大小
    hwndOwner As Long                                                           '对话框所有者的窗口句柄
    hInstance As Long                                                           '实例句柄
    lpstrFilter As String                                                       '过滤列表
    lpstrCustomFilter As String                                                 '自定义的过滤列表
    nMaxCustFilter As Long                                                      '最大自定义过滤列表
    nFilterIndex As Long                                                        '过滤列表索引
    lpstrFile As String                                                         '文件名
    nMaxFile As Long                                                            '最大文件数量
    lpstrFileTitle As String                                                    '文件标题
    nMaxFileTitle As Long                                                       '最大文件标题数量
    lpstrInitialDir As String                                                   '初始化目录
    lpstrTitle As String                                                        '对话框的标题
    flags As Long                                                               '样式
    nFileOffset As Integer                                                      '文件偏移数量
    nFileExtension As Integer                                                   '拓展名
    lpstrDefExt As String                                                       '默认拓展名
    lCustData As Long                                                           '数据
    lpfnHook As Long                                                            '函数钩子
    lpTemplateName As String                                                    '模板名称
End Type

'页面设置对话框
Private Type PageSetupDlg
    lStructSize As Long                                                         '结构体的大小
    hwndOwner As Long                                                           '对话框所有者的窗口句柄
    hDevMode As Long                                                            '设备模式句柄
    hDevNames As Long                                                           '设备名称句柄
    flags As Long                                                               '样式
    ptPaperSize As PointAPI                                                     '纸张的尺寸
    rtMinMargin As Rect                                                         '最小页边距
    rtMargin As Rect                                                            '页边距
    hInstance As Long                                                           '实例句柄
    lCustData As Long                                                           '数据
    lpfnPageSetupHook As Long                                                   '页面设置钩子
    lpfnPagePaintHook As Long                                                   '页面打印钩子
    lpPageSetupTemplateName As String                                           '页面设置模板名称
    hPageSetupTemplate As Long                                                  '页面设置临时句柄
End Type

'选择颜色对话框
Private Type ChooseColor
    lStructSize As Long                                                         '结构体的大小
    hwndOwner As Long                                                           '对话框所有者的窗口句柄
    hInstance As Long                                                           '实例句柄
    rgbResult As Long                                                           '返回结果
    lpCustColors As String                                                      '自定义颜色
    flags As Long                                                               '样式
    lCustData As Long                                                           '数据
    lpfnHook As Long                                                            '函数钩子
    lpTemplateName As String                                                    '模板名称
End Type

'逻辑字体（Logical Font）
Private Type LogFont
    lfHeight As Long                                                            '字符的高度（当lfHeight大于0时，字体映射程序将该值转换为设备单位，并将它与可用字体的字符元高度进行匹配；当该参数为0时，字体映射程度将使用一个匹配的默认高度值；如果参数的值小于0，则将其转换为设备单位，并将其绝对值与可用字体的字符高度进行匹配）
    lfWidth As Long                                                             '平均宽度（如果lfWidth的值为0，则根据设备的纵横比从可用字体的数字转换纵横中选取最接近的匹配值，该值通过比较两者之间的差异的绝对值得出）
    lfEscapement As Long                                                        '以十分之一度为单位指定每一行文本输出时相对于页面底端的角度
    lfOrientation As Long                                                       '以十分之一度为单位指定字符基线相对于页面底端的角度
    lfWeight As Long                                                            '字体重量（lfWeight的范围为0到1000，正常情况下的字体重量为400，粗体为700。如果lfWeight为0，则使用默认的字体重量）
    lfItalic As Byte                                                            '斜体
    lfUnderline As Byte                                                         '下划线
    lfStrikeOut As Byte                                                         '删除线
    lfCharSet As Byte                                                           '字符集
    lfOutPrecision As Byte                                                      '输出精度
    lfClipPrecision As Byte                                                     '剪辑精度
    lfQuality As Byte                                                           '输出质量
    lfPitchAndFamily As Byte                                                    '字体的字符间距和族
    lfFaceName As String * 31                                                   '字体名
End Type

'选择字体对话框
Private Type ChooseFont
    lStructSize As Long
    hwndOwner As Long                                                           '
    hdc As Long                                                                 '
    lpLogFont As Long                                                           '指向逻辑字体（Logical Font）的指针
    iPointSize As Long                                                          '所选字体的大小
    flags As Long                                                               '样式
    rgbColors As Long                                                           '字体的颜色
    lCustData As Long                                                           '数据
    lpfnHook As Long                                                            '函数钩子
    lpTemplateName As String                                                    '模板名称
    hInstance As Long                                                           '实例句柄
    lpszStyle As String                                                         '字体风格（必须是LF_FACESIZE或更大）
    nFontType As Integer                                                        '字体类型
    MISSING_ALIGNMENT As Integer                                                '缺失的对齐方式（暂译，功能未知）
    nSizeMin As Long                                                            '字体允许的最小尺寸
    nSizeMax As Long                                                            '字体允许的最大尺寸
End Type

'打印对话框结构
Private Type PrintDlg_Type
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long                                                                 '设备上下文（Device Contexts）句柄
    flags As Long
    nFromPage As Integer                                                        '起始页数
    nToPage As Integer                                                          '终止页数
    nMinPage As Integer                                                         '最小页数
    nMaxPage As Integer                                                         '最大页数
    nCopies As Integer                                                          '打印份数
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'设备名称结构
Private Type DevNames_Type
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

'设备模式结构
Private Type DevMode_Type
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


'返回的字体信息
'Public Type FontInfo
'
'End Type

'API声明
Private Declare Function ChooseColorDlg Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg_Type) As Long
Private Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PageSetupDlg) As Long
Private Declare Function ChooseFontDlg Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long

'自定义的变量
Dim Ofn As OpenFileName
Dim CustomColors  As String

'函数和过程

'显示“打开文件”对话框
Public Function ShowOpenDialog(Optional ByVal DialogTitle As String = "", Optional ByVal FileFilter As String = "所有文件|*.*", Optional ByVal InitialDirectory As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal OwnerHandle As Long = 0, Optional ByVal InstanceHandle As Long = 0) As String
    Const OFN_ALLOWMULTISELECT = &H200
    Const OFN_EXPLORER = &H80000
    Const OFN_FILEMUSTEXIST = &H1000
    Dim Ofn As OpenFileName
    Dim lRet As Long
    Dim Files As Variant, f As String, I As Long
    Ofn.lStructSize = Len(Ofn)
    Ofn.hwndOwner = OwnerHandle
    Ofn.hInstance = InstanceHandle
    Ofn.lpstrFilter = Replace(FileFilter, "|", Chr$(0))
    Ofn.lpstrFile = Space$(MAXFILECOUNT - 1)
    Ofn.nMaxFile = MAXFILECOUNT
    Ofn.lpstrFileTitle = Space$(MAXFILECOUNT - 1)
    Ofn.nMaxFileTitle = MAXFILECOUNT
    Ofn.lpstrInitialDir = InitialDirectory
    Ofn.lpstrTitle = IIf(DialogTitle = "", " ", DialogTitle)
    Ofn.flags = IIf(MultiSelect, OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST, OFN_EXPLORER Or OFN_FILEMUSTEXIST)
    lRet = GetOpenFileName(Ofn)
    If lRet = 0 Then
        ShowOpenDialog = ""
    Else
        If MultiSelect Then
            Files = Split(Ofn.lpstrFile, Chr$(0))
            For I = 1 To UBound(Files)
                If Trim(Files(I)) <> "" Then f = f & Files(0) & "\" & Files(I) & vbCrLf
            Next I
            If Trim(f) = "" Then
                If Dir(Files(0)) <> "" Then f = Files(0)
            Else
                f = Left(f, Len(f) - 1)
            End If
            ShowOpenDialog = Trim$(f)
        Else
            ShowOpenDialog = Trim$(Ofn.lpstrFile)
        End If
    End If
End Function

'显示“保存文件”对话框
Public Function ShowSaveDialog(Optional ByVal DialogTitle As String = "", Optional ByVal FileFilter As String = "所有文件|*.*", Optional ByVal InitialDirectory As String = "", Optional ByVal OwnerHandle As Long = 0, Optional ByVal InstanceHandle As Long = 0) As String
    Dim Ofn As OpenFileName
    Dim lRet As Long
    Dim Ans As Integer
    Ofn.lStructSize = Len(Ofn)
    Ofn.hwndOwner = OwnerHandle
    Ofn.hInstance = InstanceHandle
    Ofn.lpstrFilter = Replace(FileFilter, "|", Chr$(0))
    Ofn.lpstrFile = Space$(MAXFILECOUNT - 1)
    Ofn.nMaxFile = MAXFILECOUNT
    Ofn.lpstrFileTitle = Space$(MAXFILECOUNT - 1)
    Ofn.nMaxFileTitle = MAXFILECOUNT
    Ofn.lpstrInitialDir = InitialDirectory
    Ofn.lpstrTitle = IIf(DialogTitle = "", " ", DialogTitle)
    Ofn.flags = 0
ShowDlg:
    lRet = GetSaveFileName(Ofn)
    If lRet <> 0 Then
        If Dir(Ofn.lpstrFile) <> "" Then                                        '文件存在
            Ans = MsgBox(Replace(Trim$(Ofn.lpstrFileTitle), Chr$(0), "") & " 已存在。" & vbCrLf & "要替换它吗？", vbYesNo Or vbExclamation, "")
            If Ans = vbYes Then
                ShowSaveDialog = Trim$(Ofn.lpstrFile)
            Else
                GoTo ShowDlg
            End If
        End If
        ShowSaveDialog = Trim$(Ofn.lpstrFile)
    Else
        ShowSaveDialog = ""
    End If
End Function

'显示“颜色”对话框
Public Function ShowColorDialog(Optional ByVal OwnerHandle As Long = 0, Optional ByVal InstanceHandle As Long = 0) As Long
    Dim cc As ChooseColor
    Dim Custcolor(15) As Long
    Dim lReturn As Long
    cc.lStructSize = Len(cc)
    cc.hwndOwner = OwnerHandle
    cc.hInstance = InstanceHandle
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    If ChooseColorDlg(cc) <> 0 Then
        ShowColorDialog = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColorDialog = -1
    End If
End Function

Private Function MulDiv(a As Double, b As Double, c As Double) As Double
    MulDiv = a * b / c
End Function

'显示“字体”对话框
Public Function ShowFontDialog(Optional ByVal OwnerHandle As Long = 0) As StdFont
    Dim RetnFont As New StdFont
    Dim CF As ChooseFont, lFont As LogFont, hMem As Long, pMem As Long
    Dim fontName As String, retVal As Long
    lFont.lfHeight = 0                                                          '
    lFont.lfWidth = 0                                                           '
    lFont.lfEscapement = 0                                                      '
    lFont.lfOrientation = 0                                                     '
    lFont.lfWeight = FW_NORMAL                                                  '
    lFont.lfCharSet = DEFAULT_CHARSET                                           '
    lFont.lfOutPrecision = OUT_DEFAULT_PRECIS                                   '
    lFont.lfClipPrecision = CLIP_DEFAULT_PRECIS                                 '
    lFont.lfQuality = DEFAULT_QUALITY                                           '
    lFont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN                          '
    lFont.lfFaceName = vbNullChar                                               '
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lFont))
    pMem = GlobalLock(hMem)                                                     '
    CopyMemory ByVal pMem, lFont, Len(lFont)                                    '
    CF.lStructSize = Len(CF)                                                    '
    CF.hwndOwner = OwnerHandle                                                  '
    CF.hdc = Printer.hdc                                                        '
    CF.lpLogFont = pMem                                                         '
    CF.iPointSize = 120                                                         ' 12点字体(以1/10点为单位)
    CF.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT ' Or CF_LIMITSIZE
    CF.rgbColors = RGB(0, 0, 0)                                                 ' 黑色默认
    CF.nFontType = REGULAR_FONTTYPE                                             ' 默认字体样式
    CF.nSizeMin = 1                                                             '
    CF.nSizeMax = 1048575                                                       '
    ' 调用函数，如果成功，将LOGFONT结构复制回该结构，然后打印用户选择的属性。
    retVal = ChooseFontDlg(CF)                                                  '
    If retVal <> 0 Then                                                         '
        CopyMemory lFont, ByVal pMem, Len(lFont)                                '
        With RetnFont
            .Name = Left(lFont.lfFaceName, InStr(lFont.lfFaceName, Chr$(0)) - 1)
            .Size = Int(MulDiv(-lFont.lfHeight, 72, GetDeviceCaps(CreateDC("DISPLAY", vbNullString, vbNullString, 0), LOGPIXELSY)))
            .Charset = lFont.lfCharSet
            .Bold = lFont.lfWeight <> 0
            .Italic = lFont.lfItalic
            .Strikethrough = lFont.lfStrikeOut
            .Underline = lFont.lfUnderline
            .Weight = lFont.lfWeight
        End With
        Set ShowFontDialog = RetnFont
    Else
        With RetnFont
            .Name = ""
            '            .Size = 9
            '            .Charset = 0
            '            .Bold = False
            '            .Italic = False
            '            .Strikethrough = False
            '            .Underline = False
            '            .Weight = 400
        End With
        Set ShowFontDialog = RetnFont
    End If
    retVal = GlobalUnlock(hMem)                                                 '销毁指针，解锁内存块
    retVal = GlobalFree(hMem)                                                   ' 释放已分配的内存
    Set RetnFont = Nothing
End Function

Private Sub Class_Initialize()
    CustomColors = String$(32, Chr$(0))                                         '自定义颜色的16个格子
End Sub
