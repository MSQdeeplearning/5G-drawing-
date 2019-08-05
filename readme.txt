requirement:
pillow
baidu-aip
numpy
re
glob
os
base64
requests
time
opencv
xlwt

1. 第一个框选的必须是表格，第二个框和第三个框顺序随意。
2. 识别文字效果拔群，但是个位数字的识别无法通过调用百度OCR完成
3. 将py文件放在和识别的图片同一目录下即可直接运行程序
4. 识别完成后，表格会以excel的形式存储在当前目录下，另外两个以txt的格式存储在当前目录下，名称均为识别图片的名称。
