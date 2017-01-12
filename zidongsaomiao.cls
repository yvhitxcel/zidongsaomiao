Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'版本：1.0.0
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
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Private Type POINT
    x As Long
    Y As Long
End Type

Private Sub SingleClick() '防止屏幕黑屏，模拟鼠标单击
  Dim pLocation As POINT
  Call GetCursorPos(pLocation)
  SetCursorPos pLocation.x, pLocation.Y 'x and y position
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
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
    '    '245 331 250 1000 1040  1040-245=795/210=3.7857  210-23-10=177*3.7587=665
    '    '23*3.7587=86.4501+245=331.45
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

Sub turn_OLD() '没用了，也算是一个思路吧，留着
    '是个笨办法，不算角度，判断文件有偏就转一定的角度，直到文件端正。
    Dim lDC As Variant
    Dim j As Integer, R As Integer
    Dim lngCor As Long
    Dim L1 As Integer, L2 As Integer
    
    lDC = GetWindowDC(0)
    For j = 250 To 400
        lngCor = GetPixel(lDC, j, 560)
        R = lngCor Mod 256
        If R < 200 Then
            L1 = j
            Exit For
        End If
    Next
    For j = 250 To 400
        lngCor = GetPixel(lDC, j, 720)
        R = lngCor Mod 256
        If R < 200 Then
            L2 = j
            Exit For
        End If
    Next
    If (L2 > L1 And L2 > 0 And L1 > 0 And L2 < 400 And L1 < 400) Then
        ActiveDocument.Shapes(1).IncrementRotation 0.1 '向左
        s2
        turn_OLD
    End If
    If (L2 < L1 And L2 > 0 And L1 > 0 And L2 < 400 And L1 < 400) Then
        ActiveDocument.Shapes(1).IncrementRotation -0.05 '向右
        s2
        turn_OLD
    End If
End Sub
