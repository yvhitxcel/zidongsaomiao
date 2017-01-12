# zidongsaomiao
自动扫描图片，对图片进行缩放，偏移，旋转操作，以满足竣工资料的规范要求。

# 使用说明
打开WORD，点击开发工具，启动Visual Basic编辑器，复制此代码到代码窗口，点击begging，按F5运行。
前提条件是满足以下步骤。

# 步骤
## 1、设置屏幕分辨率为 1280*800
国为代码的各个参数是在这个分辨率下采集来的，当然你也可以使用其它分辨率，但是参数就得你自己调整了。

## 2、准备目录
在D盘新建 zidongsaomiao 文件夹，并在这个目录下，再新建input和ouuput文件夹。

## 3、准备好预加载图片
复制项目内297-1.png图片到 D:\zidongsaomiao\ 文件夹
图片大小为 宽1mm,高297mm，必需在加载文件图片之前就加载到文档，作用就是促使窗口跳转到最底部。

## 4、准备好需要处理的文件图片
使用Adobe软件打xxx.pdf文档，点击文件另存为jpeg图片，保存位置为第二步创建的 d:\zidongsaomiao\input\ 目录。
可以观察到图片命名规则为 xxx_页面_nnn.jpg，如果pdf文件少于100页，那你需要在代码处改一下规则，否则会遇到找不到文件的错误。

## 5、打开WORD软件，点击文件另存为，检查是否能另存为PDF文档
代码处理完成之后默认是另存为PDF文档，因此需要用到ADOBE的这个插件。

## 6、结束
在代码窗口打到并运行 begging ，处理速度约为 5页/分钟。
当运行结束之后，就能在output目录下见到处理过后的PDF文件啦。
