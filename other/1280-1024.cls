'’Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'版本：1.0.3
'作者：唐海勇
' Q Q：2338953
'邮箱：yvhitxcel@tom.com
'微信：2338953
'网址：http://github.com/yvhitxcel/zidongsaomiao
'协议：GNU GENERAL PUBLIC LICENSE Version 3.0
'
'说明：开发源自于实际工作，项目需要处理大量PDF文件，文件由扫描仪批量扫描生成， _
'       由于原始文件来源非常广泛，各种宽度杂，另外扫描仪扫描时也导致文件的偏斜， _
'       导致目前的工作进度缓慢，此前一直由人工的办法，对文件进行逐页修改，人均 _
'       每天处理量约为150页，项目截止到目前大约还有近50万页文件待处理，任务非常艰巨。
'思路：先把pdf文件另存为jpeg文件格式，然后利用word宏语言，动态加载图片到WORD文档 _
'       经过屏幕取色函数判断文件边框所在位置，进而依据此数据进行，缩放，旋转，平 _
'       移操作,达到印刷工业标准需求。
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
    Dim strFilename As String   '图片规则为 "xxx_页面_xxx.jpg",当PDF文件另存为jpeg时可自动生成
    Dim FileName As String      '文件名前缀 "xxx_页面_"
    Dim strFilenameSave As String
    Dim im As Shape, imm As InlineShape
    Dim strMulu As String
    Dim intStartPage As Integer '在此配置图片页码开始值
    Dim intStoopPage As Integer '在此配置图片页码结束值
    
    intStartPage = 1
    intStoopPage = 353
    FileName = "C4RX 土建结构工程 VOL-3_页面_"
    
    strMulu = "D:\zidongsaomiao"  '工作目录，目录下有input output两个文件夹
                                  ' input内为原始图片
                                  'output为输出文件夹
    For i = intStartPage To intStoopPage
        strFilename = FileName
        If (i < 10) Then strFilename = strFilename & "00" & i
        If (i > 9 And i < 100) Then strFilename = strFilename & "0" & i
        If (i > 99) Then strFilename = strFilename & i
        strFilenameSave = strFilename & ".pdf"      '处理完成后另存为PDF格式
        strFilename = strFilename & ".jpg"          '初始文件名

        '打开需要处理的文件  '暂停时间需要长达3秒
        'Selection.InlineShapes.AddPicture FileName:= _
            strMulu & "\input\" & strFilename _
            , LinkToFile:=False, SaveWithDocument:=True
        'For Each imm In ActiveDocument.InlineShapes
        '    Events 1
        '    imm.ConvertToShape '图片插入时为嵌入型，修改成浮于文件上方，这样才移动旋转。
        '    Events 2
        'Next
        
        '以shapes方式直接加载图片，可节约图片转换所花时间
        Selection.InlineShapes.AddPicture FileName:= _
            strMulu & "\297-1.png" _
            , LinkToFile:=False, SaveWithDocument:=True  '打开一个297mm*1mm规格空白png图片，促使页面停靠在窗口底端
        Events 0.3    '等待文件停靠到窗口最底端
       
        ActiveDocument.Shapes.AddPicture FileName:= _
            strMulu & "\input\" & strFilename _
            , LinkToFile:=False, SaveWithDocument:=True
        Events 1      '等待图片加载完成

        worknum = 0   '重复运行worker，不需要暂停
        worker        '进入加工流程
        worker
        worker

        '另存为PDF文件
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            strMulu & "\output\" & strFilenameSave, ExportFormat:= _
            wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
            Item:=wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=False, _
            CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=False, UseISO19005_1:=False
        For Each im In ActiveDocument.Shapes
            im.Delete   '删除所有图片
        Next
        For Each imm In ActiveDocument.InlineShapes
            imm.Delete  '删除所有图片
        Next
        SingleClick  '防止进入屏保
    Next
    
End Sub

Sub worker() '图片处理工人
    Dim Xtmp As Integer, boolX As Integer
    Dim j As Integer
    Dim H As Long
    Dim x, Yarea1, Yarea2, Xarea1, Xarea2
    Dim lDC As Variant
    Dim Mnub As Double, dushu As Double
    Dim bili As Double
    Dim X2 As Integer, X1 As Integer
    Dim Y2 As Integer, Y1 As Integer, Yn As Integer
    Dim L1 As Integer, R1 As Integer  '最佳位置左右边距
    Dim lngCor As Long
    Dim n As Integer       '偶尔会有因延时不够导致取不到边框的情况，重复三次

    lDC = GetWindowDC(0)
    '页面默认显示下半端
    n = 0
    While (X2 = 0 And n <= 2)
        Yarea1 = 190    '页面显示下半端时，文件最上端在屏幕上的Y抽位置185 + 5
        Yarea2 = 953    '页面显示下半端时，文件最下端在屏幕上的Y抽位置730 - 2 955，估算坚线的高度区间小于728
        Xarea1 = 250    '页面显示下半端时，文件最左端在屏幕上的X抽位置245 + 5
        Xarea2 = 400    '页面显示下半端时，文件最左端在屏幕上的X抽位置400，估算坚线的水平位置区间小于400
        Xtmp = 0
        X2 = 0
        
        While CInt(Yarea1 + 0.5) < Int(Yarea2) '垂直方向使用二分法加速查找边框端点所在位置
            Mnub = (Yarea1 + Yarea2) / 2
            boolX = 0
            For x = Xarea1 To Xarea2   '水平X抽方向直接从左向逐点查找边框所在位置
                lngCor = GetPixel(lDC, x, Mnub)
                RGBtoH lngCor, H
                If H < 190 Then
                    X2 = Xtmp        '前几次得到的边距相同，最后一个因图片模糊，导致数值突变
                    If (Xtmp > x + 1) Then
                        Mnub = Int(Yarea1) + 1
                    Else
                        Xarea1 = x - 10  '减少X抽区间，加快查找速度。
                        Xarea2 = x + 10
                        Xtmp = x         '因此增加了对上一次结果的缓存，不采用最后一次结果。
                    End If
                    If (X2 = 0) Then X2 = Xtmp
                    boolX = x
                    Exit For
                End If
            Next x
            If boolX > 0 Then
                Yarea1 = Mnub       '此处，当文件显示下半端或显示上半端时 不同
            Else
                Yarea2 = Mnub
            End If
        Wend
        X2 = X2
        Y2 = 953 - Int(Yarea1)  '页面显示下半端时，边框最底端到文件最底端的距离,文件最低端在屏幕上的Y抽位置731
        n = n + 1
        If (X2 = 0) Then Events 1
    Wend
    
    '通过下面命令可使页面显示上半端
    If (worknum = 0) Then
        Application.ActiveWindow.Selection.GoTo wdGoToPage, wdGoToNext, , 1
        Events 1
    End If
    
    n = 0
    While (X1 = 0 And n <= 2)
        Yarea1 = 176    '页面显示上半端时，文件最上端在屏幕上的Y抽位置200 + 3
        Yarea2 = 765 * 2 - 188 '页面显示上半端时，文件最下端在屏幕上的Y抽位置748
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
        Y1 = CInt(Yarea2 + 0.5) - 173   '边框最顶端到文件最顶端的距离,文件最顶端在屏幕上的位置是200
        If (Yarea2 < 560) Then Yn = 660 Else Yn = Int(Yarea2) + 100 '边框最顶端向下100px,以此位置取最左到最最右位置，判断缩放与移动
        n = n + 1
    Wend
    
    bili = 1
    dushu = 0
    If (X1 > 0 And X2 > 0 And Y1 > 0 And Y2 > 0) Then
        dushu = Atn((X2 - X1) / (1122 - Y2 - Y1)) * 180 / 3.1415926 '文件总高度1122px 已经直角三角形的两垂直边的长度，求角度公式  角度*180/pi

        For j = 1038 To 800 Step -1          '从右向左查找右边框位置
            lngCor = GetPixel(lDC, j, Yn)
            RGBtoH lngCor, H
            If H < 190 Then
                R1 = j                       '得到右边框水平位置
                Exit For
            End If
        Next

        bili = 664.5 / (R1 - X1)                                    '文件左右边框要求的宽度为664.5px(210-23-11=176*3.774)
        If (Abs(bili) > 1.2 Or Abs(bili) < 0.8) Then bili = 1
        With ActiveDocument.Shapes(1)
            If (dushu <> 0 And Abs(dushu) < 1 And worknum = 0) Then
                .IncrementRotation dushu                            '正数X2>X1 顺时针转 向右
            End If
            .Left = (332 - (((X1 + X2) / 2) - .Left)) * bili        '实施移动操作 左边框要求的位置为332px 23*3.774+245
            If (bili <> 1) Then
                .ScaleWidth bili, False                             '设置高度
                .ScaleHeight bili, False                            '设置宽度
            End If
        End With
        If (worknum = 0) Then Events 1                              'n秒后进行保存操作
        worknum = worknum + 1
    End If
End Sub

Sub Events(n As Double) '暂停n秒，期间可以进行其他操作,主要用于等待页面加载完成
    Dim t As Double
    t = Timer
    While Timer < t + n
        DoEvents
    Wend
End Sub

Private Sub SingleClick() '防止屏幕黑屏，模拟鼠标单击
  Dim pLocation As POINT
  Call GetCursorPos(pLocation)
  SetCursorPos pLocation.x, pLocation.Y 'x and y position
  mouse_event &H2, 0, 0, 0, 0
  mouse_event &H4, 0, 0, 0, 0
End Sub

Sub RGBtoH(lColour As Long, H As Long) 'RGB转换成灰度值
    Dim R As Long, G As Long, B As Long
    R = lColour Mod 256
    G = ((lColour And &HFF00&) \ 256&) Mod 256&
    B = (lColour And &HFF0000) \ 65536
    H = (R * 77 + G * 150 + B * 29) \ 256 '三原色换成灰度值公式
End Sub

Sub get_point() '辅助功能
    '此功能用于查找各个位置在屏幕的值，得天值之后配置到上面各个相应位置。
    Dim pLocation As POINT
    Dim lColour As Long
    Dim lDC As Variant, R, G, B
    'Application.ActiveWindow.Selection.GoTo wdGoToPage, wdGoToNext, , 1
    lDC = GetWindowDC(0)
    Call GetCursorPos(pLocation)
    lColour = GetPixel(lDC, pLocation.x, pLocation.Y)   '需要的坐标值在这里。
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

