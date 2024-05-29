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

'ͨ�öԻ���
'����ģ������comdlg32.dll���
'�ı������ߣ�Donald Grover��Դ���롣
'�޸ģ����ư��侩��

'�汾��0.1

'����
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
'�ṹ��

'��
Private Type PointAPI
    X As Long
    Y As Long
End Type

'����
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'���ļ��Ի���
Private Type OpenFileName
    lStructSize As Long                                                         '�ṹ��Ĵ�С
    hwndOwner As Long                                                           '�Ի��������ߵĴ��ھ��
    hInstance As Long                                                           'ʵ�����
    lpstrFilter As String                                                       '�����б�
    lpstrCustomFilter As String                                                 '�Զ���Ĺ����б�
    nMaxCustFilter As Long                                                      '����Զ�������б�
    nFilterIndex As Long                                                        '�����б�����
    lpstrFile As String                                                         '�ļ���
    nMaxFile As Long                                                            '����ļ�����
    lpstrFileTitle As String                                                    '�ļ�����
    nMaxFileTitle As Long                                                       '����ļ���������
    lpstrInitialDir As String                                                   '��ʼ��Ŀ¼
    lpstrTitle As String                                                        '�Ի���ı���
    flags As Long                                                               '��ʽ
    nFileOffset As Integer                                                      '�ļ�ƫ������
    nFileExtension As Integer                                                   '��չ��
    lpstrDefExt As String                                                       'Ĭ����չ��
    lCustData As Long                                                           '����
    lpfnHook As Long                                                            '��������
    lpTemplateName As String                                                    'ģ������
End Type

'ҳ�����öԻ���
Private Type PageSetupDlg
    lStructSize As Long                                                         '�ṹ��Ĵ�С
    hwndOwner As Long                                                           '�Ի��������ߵĴ��ھ��
    hDevMode As Long                                                            '�豸ģʽ���
    hDevNames As Long                                                           '�豸���ƾ��
    flags As Long                                                               '��ʽ
    ptPaperSize As PointAPI                                                     'ֽ�ŵĳߴ�
    rtMinMargin As Rect                                                         '��Сҳ�߾�
    rtMargin As Rect                                                            'ҳ�߾�
    hInstance As Long                                                           'ʵ�����
    lCustData As Long                                                           '����
    lpfnPageSetupHook As Long                                                   'ҳ�����ù���
    lpfnPagePaintHook As Long                                                   'ҳ���ӡ����
    lpPageSetupTemplateName As String                                           'ҳ������ģ������
    hPageSetupTemplate As Long                                                  'ҳ��������ʱ���
End Type

'ѡ����ɫ�Ի���
Private Type ChooseColor
    lStructSize As Long                                                         '�ṹ��Ĵ�С
    hwndOwner As Long                                                           '�Ի��������ߵĴ��ھ��
    hInstance As Long                                                           'ʵ�����
    rgbResult As Long                                                           '���ؽ��
    lpCustColors As String                                                      '�Զ�����ɫ
    flags As Long                                                               '��ʽ
    lCustData As Long                                                           '����
    lpfnHook As Long                                                            '��������
    lpTemplateName As String                                                    'ģ������
End Type

'�߼����壨Logical Font��
Private Type LogFont
    lfHeight As Long                                                            '�ַ��ĸ߶ȣ���lfHeight����0ʱ������ӳ����򽫸�ֵת��Ϊ�豸��λ�������������������ַ�Ԫ�߶Ƚ���ƥ�䣻���ò���Ϊ0ʱ������ӳ��̶Ƚ�ʹ��һ��ƥ���Ĭ�ϸ߶�ֵ�����������ֵС��0������ת��Ϊ�豸��λ�����������ֵ�����������ַ��߶Ƚ���ƥ�䣩
    lfWidth As Long                                                             'ƽ�����ȣ����lfWidth��ֵΪ0��������豸���ݺ�ȴӿ������������ת���ݺ���ѡȡ��ӽ���ƥ��ֵ����ֵͨ���Ƚ�����֮��Ĳ���ľ���ֵ�ó���
    lfEscapement As Long                                                        '��ʮ��֮һ��Ϊ��λָ��ÿһ���ı����ʱ�����ҳ��׶˵ĽǶ�
    lfOrientation As Long                                                       '��ʮ��֮һ��Ϊ��λָ���ַ����������ҳ��׶˵ĽǶ�
    lfWeight As Long                                                            '����������lfWeight�ķ�ΧΪ0��1000����������µ���������Ϊ400������Ϊ700�����lfWeightΪ0����ʹ��Ĭ�ϵ�����������
    lfItalic As Byte                                                            'б��
    lfUnderline As Byte                                                         '�»���
    lfStrikeOut As Byte                                                         'ɾ����
    lfCharSet As Byte                                                           '�ַ���
    lfOutPrecision As Byte                                                      '�������
    lfClipPrecision As Byte                                                     '��������
    lfQuality As Byte                                                           '�������
    lfPitchAndFamily As Byte                                                    '������ַ�������
    lfFaceName As String * 31                                                   '������
End Type

'ѡ������Ի���
Private Type ChooseFont
    lStructSize As Long
    hwndOwner As Long                                                           '
    hdc As Long                                                                 '
    lpLogFont As Long                                                           'ָ���߼����壨Logical Font����ָ��
    iPointSize As Long                                                          '��ѡ����Ĵ�С
    flags As Long                                                               '��ʽ
    rgbColors As Long                                                           '�������ɫ
    lCustData As Long                                                           '����
    lpfnHook As Long                                                            '��������
    lpTemplateName As String                                                    'ģ������
    hInstance As Long                                                           'ʵ�����
    lpszStyle As String                                                         '�����񣨱�����LF_FACESIZE�����
    nFontType As Integer                                                        '��������
    MISSING_ALIGNMENT As Integer                                                'ȱʧ�Ķ��뷽ʽ�����룬����δ֪��
    nSizeMin As Long                                                            '������������С�ߴ�
    nSizeMax As Long                                                            '�������������ߴ�
End Type

'��ӡ�Ի���ṹ
Private Type PrintDlg_Type
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long                                                                 '�豸�����ģ�Device Contexts�����
    flags As Long
    nFromPage As Integer                                                        '��ʼҳ��
    nToPage As Integer                                                          '��ֹҳ��
    nMinPage As Integer                                                         '��Сҳ��
    nMaxPage As Integer                                                         '���ҳ��
    nCopies As Integer                                                          '��ӡ����
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'�豸���ƽṹ
Private Type DevNames_Type
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

'�豸ģʽ�ṹ
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


'���ص�������Ϣ
'Public Type FontInfo
'
'End Type

'API����
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

'�Զ���ı���
Dim Ofn As OpenFileName
Dim CustomColors  As String

'�����͹���

'��ʾ�����ļ����Ի���
Public Function ShowOpenDialog(Optional ByVal DialogTitle As String = "", Optional ByVal FileFilter As String = "�����ļ�|*.*", Optional ByVal InitialDirectory As String = "", Optional ByVal MultiSelect As Boolean = False, Optional ByVal OwnerHandle As Long = 0, Optional ByVal InstanceHandle As Long = 0) As String
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

'��ʾ�������ļ����Ի���
Public Function ShowSaveDialog(Optional ByVal DialogTitle As String = "", Optional ByVal FileFilter As String = "�����ļ�|*.*", Optional ByVal InitialDirectory As String = "", Optional ByVal OwnerHandle As Long = 0, Optional ByVal InstanceHandle As Long = 0) As String
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
        If Dir(Ofn.lpstrFile) <> "" Then                                        '�ļ�����
            Ans = MsgBox(Replace(Trim$(Ofn.lpstrFileTitle), Chr$(0), "") & " �Ѵ��ڡ�" & vbCrLf & "Ҫ�滻����", vbYesNo Or vbExclamation, "")
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

'��ʾ����ɫ���Ի���
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

'��ʾ�����塱�Ի���
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
    CF.iPointSize = 120                                                         ' 12������(��1/10��Ϊ��λ)
    CF.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT ' Or CF_LIMITSIZE
    CF.rgbColors = RGB(0, 0, 0)                                                 ' ��ɫĬ��
    CF.nFontType = REGULAR_FONTTYPE                                             ' Ĭ��������ʽ
    CF.nSizeMin = 1                                                             '
    CF.nSizeMax = 1048575                                                       '
    ' ���ú���������ɹ�����LOGFONT�ṹ���ƻظýṹ��Ȼ���ӡ�û�ѡ������ԡ�
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
    retVal = GlobalUnlock(hMem)                                                 '����ָ�룬�����ڴ��
    retVal = GlobalFree(hMem)                                                   ' �ͷ��ѷ�����ڴ�
    Set RetnFont = Nothing
End Function

Private Sub Class_Initialize()
    CustomColors = String$(32, Chr$(0))                                         '�Զ�����ɫ��16������
End Sub