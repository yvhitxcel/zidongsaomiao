'��Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�汾��1.0.3
'���ߣ��ƺ���
' Q Q��2338953
'���䣺yvhitxcel@tom.com
'΢�ţ�2338953
'��ַ��http://github.com/yvhitxcel/zidongsaomiao
'Э�飺GNU GENERAL PUBLIC LICENSE Version 3.0
'
'˵��������Դ����ʵ�ʹ�������Ŀ��Ҫ��������PDF�ļ����ļ���ɨ��������ɨ�����ɣ� _
'       ����ԭʼ�ļ���Դ�ǳ��㷺�����ֿ����ӣ�����ɨ����ɨ��ʱҲ�����ļ���ƫб�� _
'       ����Ŀǰ�Ĺ������Ȼ�������ǰһֱ���˹��İ취�����ļ�������ҳ�޸ģ��˾� _
'       ÿ�촦����ԼΪ150ҳ����Ŀ��ֹ��Ŀǰ��Լ���н�50��ҳ�ļ�������������ǳ���ޡ�
'˼·���Ȱ�pdf�ļ�����Ϊjpeg�ļ���ʽ��Ȼ������word�����ԣ���̬����ͼƬ��WORD�ĵ� _
'       ������Ļȡɫ�����ж��ļ��߿�����λ�ã��������ݴ����ݽ��У����ţ���ת��ƽ _
'       �Ʋ���,�ﵽӡˢ��ҵ��׼����
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CancelDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Private Type POINT
    x As Long
    Y As Long
End Type
Dim worknum As Integer
                        
Sub begging()
    Dim i As Integer
    Dim strFilename As String   'ͼƬ����Ϊ "xxx_ҳ��_xxx.jpg",��PDF�ļ�����Ϊjpegʱ���Զ�����
    Dim FileName As String      '�ļ���ǰ׺ "xxx_ҳ��_"
    Dim strFilenameSave As String
    Dim im As Shape, imm As InlineShape
    Dim strMulu As String
    Dim intStartPage As Integer '�ڴ�����ͼƬҳ�뿪ʼֵ
    Dim intStoopPage As Integer '�ڴ�����ͼƬҳ�����ֵ
    
    intStartPage = 1
    intStoopPage = 353
    FileName = "C4RX �����ṹ���� VOL-3_ҳ��_"
    
    strMulu = "D:\zidongsaomiao"  '����Ŀ¼��Ŀ¼����input output�����ļ���
                                  ' input��ΪԭʼͼƬ
                                  'outputΪ����ļ���
    For i = intStartPage To intStoopPage
        strFilename = FileName
        If (i < 10) Then strFilename = strFilename & "00" & i
        If (i > 9 And i < 100) Then strFilename = strFilename & "0" & i
        If (i > 99) Then strFilename = strFilename & i
        strFilenameSave = strFilename & ".pdf"      '������ɺ�����ΪPDF��ʽ
        strFilename = strFilename & ".jpg"          '��ʼ�ļ���

        '����Ҫ�������ļ�  '��ͣʱ����Ҫ����3��
        'Selection.InlineShapes.AddPicture FileName:= _
            strMulu & "\input\" & strFilename _
            , LinkToFile:=False, SaveWithDocument:=True
        'For Each imm In ActiveDocument.InlineShapes
        '    Events 1
        '    imm.ConvertToShape 'ͼƬ����ʱΪǶ���ͣ��޸ĳɸ����ļ��Ϸ����������ƶ���ת��
        '    Events 2
        'Next
        
        '��shapes��ʽֱ�Ӽ���ͼƬ���ɽ�ԼͼƬת������ʱ��
        Selection.InlineShapes.AddPicture FileName:= _
            strMulu & "\297-1.png" _
            , LinkToFile:=False, SaveWithDocument:=True  '��һ��297mm*1mm���հ�pngͼƬ����ʹҳ��ͣ���ڴ��ڵ׶�
        Events 0.3    '�ȴ��ļ�ͣ����������׶�
       
        ActiveDocument.Shapes.AddPicture FileName:= _
            strMulu & "\input\" & strFilename _
            , LinkToFile:=False, SaveWithDocument:=True
        Events 1      '�ȴ�ͼƬ�������

        worknum = 0   '�ظ�����worker������Ҫ��ͣ
        worker        '����ӹ�����
        worker
        worker

        '����ΪPDF�ļ�
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            strMulu & "\output\" & strFilenameSave, ExportFormat:= _
            wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
            Item:=wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=False, _
            CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=False, UseISO19005_1:=False
        For Each im In ActiveDocument.Shapes
            im.Delete   'ɾ������ͼƬ
        Next
        For Each imm In ActiveDocument.InlineShapes
            imm.Delete  'ɾ������ͼƬ
        Next
        SingleClick  '��ֹ��������
    Next
    
End Sub

Sub worker() 'ͼƬ��������
    Dim Xtmp As Integer, boolX As Integer
    Dim j As Integer
    Dim H As Long
    Dim x, Yarea1, Yarea2, Xarea1, Xarea2
    Dim lDC As Variant
    Dim Mnub As Double, dushu As Double
    Dim bili As Double
    Dim X2 As Integer, X1 As Integer
    Dim Y2 As Integer, Y1 As Integer, Yn As Integer
    Dim L1 As Integer, R1 As Integer  '���λ�����ұ߾�
    Dim lngCor As Long
    Dim n As Integer       'ż����������ʱ��������ȡ�����߿��������ظ�����

    lDC = GetWindowDC(0)
    'ҳ��Ĭ����ʾ�°��
    n = 0
    While (X2 = 0 And n <= 2)
        Yarea1 = 190    'ҳ����ʾ�°��ʱ���ļ����϶�����Ļ�ϵ�Y��λ��185 + 5
        Yarea2 = 953    'ҳ����ʾ�°��ʱ���ļ����¶�����Ļ�ϵ�Y��λ��730 - 2 955��������ߵĸ߶�����С��728
        Xarea1 = 250    'ҳ����ʾ�°��ʱ���ļ����������Ļ�ϵ�X��λ��245 + 5
        Xarea2 = 400    'ҳ����ʾ�°��ʱ���ļ����������Ļ�ϵ�X��λ��400��������ߵ�ˮƽλ������С��400
        Xtmp = 0
        X2 = 0
        
        While CInt(Yarea1 + 0.5) < Int(Yarea2) '��ֱ����ʹ�ö��ַ����ٲ��ұ߿�˵�����λ��
            Mnub = (Yarea1 + Yarea2) / 2
            boolX = 0
            For x = Xarea1 To Xarea2   'ˮƽX�鷽��ֱ�Ӵ����������ұ߿�����λ��
                lngCor = GetPixel(lDC, x, Mnub)
                RGBtoH lngCor, H
                If H < 190 Then
                    X2 = Xtmp        'ǰ���εõ��ı߾���ͬ�����һ����ͼƬģ����������ֵͻ��
                    If (Xtmp > x + 1) Then
                        Mnub = Int(Yarea1) + 1
                    Else
                        Xarea1 = x - 10  '����X�����䣬�ӿ�����ٶȡ�
                        Xarea2 = x + 10
                        Xtmp = x         '��������˶���һ�ν���Ļ��棬���������һ�ν����
                    End If
                    If (X2 = 0) Then X2 = Xtmp
                    boolX = x
                    Exit For
                End If
            Next x
            If boolX > 0 Then
                Yarea1 = Mnub       '�˴������ļ���ʾ�°�˻���ʾ�ϰ��ʱ ��ͬ
            Else
                Yarea2 = Mnub
            End If
        Wend
        X2 = X2
        Y2 = 953 - Int(Yarea1)  'ҳ����ʾ�°��ʱ���߿���׶˵��ļ���׶˵ľ���,�ļ���Ͷ�����Ļ�ϵ�Y��λ��731
        n = n + 1
        If (X2 = 0) Then Events 1
    Wend
    
    'ͨ�����������ʹҳ����ʾ�ϰ��
    If (worknum = 0) Then
        Application.ActiveWindow.Selection.GoTo wdGoToPage, wdGoToNext, , 1
        Events 1
    End If
    
    n = 0
    While (X1 = 0 And n <= 2)
        Yarea1 = 176    'ҳ����ʾ�ϰ��ʱ���ļ����϶�����Ļ�ϵ�Y��λ��200 + 3
        Yarea2 = 765 * 2 - 188 'ҳ����ʾ�ϰ��ʱ���ļ����¶�����Ļ�ϵ�Y��λ��748
        Xtmp = 0
        X1 = 0
        While Int(Yarea1) < Int(Yarea2 - 0.5)
            Mnub = Int((Yarea1 + Yarea2) / 2)
            boolX = 0
            For x = Xarea1 To Xarea2
                lngCor = GetPixel(lDC, x, Mnub)
                RGBtoH lngCor, H
                If H < 190 Then
                    X1 = Xtmp
                    If (Xtmp < x And Xtmp > 0) Then
                        Mnub = Int(Yarea2) - 1
                    Else
                        Xarea1 = x - 10
                        Xarea2 = x + 10
                        Xtmp = x
                    End If
                    boolX = x
                    If (X1 = 0) Then X1 = Xtmp
                    Exit For
                End If
            Next x
            If boolX > 0 Then
                Yarea2 = Mnub
            Else
                Yarea1 = Mnub
            End If
        Wend
        X1 = X1
        Y1 = CInt(Yarea2 + 0.5) - 173   '�߿���˵��ļ���˵ľ���,�ļ��������Ļ�ϵ�λ����200
        If (Yarea2 < 560) Then Yn = 660 Else Yn = Int(Yarea2) + 100 '�߿��������100px,�Դ�λ��ȡ����������λ�ã��ж��������ƶ�
        n = n + 1
    Wend
    
    bili = 1
    dushu = 0
    If (X1 > 0 And X2 > 0 And Y1 > 0 And Y2 > 0) Then
        dushu = Atn((X2 - X1) / (1122 - Y2 - Y1)) * 180 / 3.1415926 '�ļ��ܸ߶�1122px �Ѿ�ֱ�������ε�����ֱ�ߵĳ��ȣ���Ƕȹ�ʽ  �Ƕ�*180/pi

        For j = 1038 To 800 Step -1          '������������ұ߿�λ��
            lngCor = GetPixel(lDC, j, Yn)
            RGBtoH lngCor, H
            If H < 190 Then
                R1 = j                       '�õ��ұ߿�ˮƽλ��
                Exit For
            End If
        Next

        bili = 664.5 / (R1 - X1)                                    '�ļ����ұ߿�Ҫ��Ŀ���Ϊ664.5px(210-23-11=176*3.774)
        If (Abs(bili) > 1.2 Or Abs(bili) < 0.8) Then bili = 1
        With ActiveDocument.Shapes(1)
            If (dushu <> 0 And Abs(dushu) < 1 And worknum = 0) Then
                .IncrementRotation dushu                            '����X2>X1 ˳ʱ��ת ����
            End If
            .Left = (332 - (((X1 + X2) / 2) - .Left)) * bili        'ʵʩ�ƶ����� ��߿�Ҫ���λ��Ϊ332px 23*3.774+245
            If (bili <> 1) Then
                .ScaleWidth bili, False                             '���ø߶�
                .ScaleHeight bili, False                            '���ÿ���
            End If
        End With
        If (worknum = 0) Then Events 1                              'n�����б������
        worknum = worknum + 1
    End If
End Sub

Sub Events(n As Double) '��ͣn�룬�ڼ���Խ�����������,��Ҫ���ڵȴ�ҳ��������
    Dim t As Double
    t = Timer
    While Timer < t + n
        DoEvents
    Wend
End Sub

Private Sub SingleClick() '��ֹ��Ļ������ģ����굥��
  Dim pLocation As POINT
  Call GetCursorPos(pLocation)
  SetCursorPos pLocation.x, pLocation.Y 'x and y position
  mouse_event &H2, 0, 0, 0, 0
  mouse_event &H4, 0, 0, 0, 0
End Sub

Sub RGBtoH(lColour As Long, H As Long) 'RGBת���ɻҶ�ֵ
    Dim R As Long, G As Long, B As Long
    R = lColour Mod 256
    G = ((lColour And &HFF00&) \ 256&) Mod 256&
    B = (lColour And &HFF0000) \ 65536
    H = (R * 77 + G * 150 + B * 29) \ 256 '��ԭɫ���ɻҶ�ֵ��ʽ
End Sub

Sub get_point() '��������
    '�˹������ڲ��Ҹ���λ������Ļ��ֵ������ֵ֮�����õ����������Ӧλ�á�
    Dim pLocation As POINT
    Dim lColour As Long
    Dim lDC As Variant, R, G, B
    'Application.ActiveWindow.Selection.GoTo wdGoToPage, wdGoToNext, , 1
    lDC = GetWindowDC(0)
    Call GetCursorPos(pLocation)
    lColour = GetPixel(lDC, pLocation.x, pLocation.Y)   '��Ҫ������ֵ�����
    '    '(746-185)*2=561*2= 1122 /297=3.777777
    'm = pLocation.x
    'n = pLocation.y
    'kuandu = PointsToMillimeters(pLocation.x)
    '    '86.43 366.8889
    '    'ActiveDocument.InlineShapes(n).Height = MillimetersToPoints(297)
    '    'aaa = PointsToMillimeters(ActiveDocument.InlineShapes(n).Width)
    R = lColour Mod 256
    G = ((lColour And &HFF00&) \ 256&) Mod 256&
    B = (lColour And &HFF0000) \ 65536
End Sub
