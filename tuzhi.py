import PIL.Image
from PIL import ImageTk
from tkinter import *    
import tkinter
from aip import AipOcr
import re
import numpy as np
import glob
import os
import base64
import requests
import time
import cv2
import xlwt  

class MainApp(Frame):
    def __init__(self,master):
        Frame.__init__(self,master=None)
        self.x = self.y = 0
        self.pos = []
        self.canvas = Canvas(self, width = 1200, height = 800, cursor="cross")
        
        self.sbarv=Scrollbar(self,orient=VERTICAL)
        self.sbarh=Scrollbar(self,orient=HORIZONTAL)
        self.sbarv.config(command=self.canvas.yview)
        self.sbarh.config(command=self.canvas.xview)

        self.canvas.config(yscrollcommand=self.sbarv.set)
        self.canvas.config(xscrollcommand=self.sbarh.set)

        self.canvas.grid(row=0,column=0,sticky=N+S+E+W)
        self.sbarv.grid(row=0,column=1,stick=N+S)
        self.sbarh.grid(row=1,column=0,sticky=E+W)
        
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_move_press)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        
        self.rect = None
        self.start_x = None
        self.start_y = None
        self.num=0                 #the number of image
        self.url = img_path[self.num]

        
        self.im = PIL.Image.open(self.url)
        self.wazil,self.lard = self.im.size
        #self.wazil,self.lard = int(self.wazil/3),int(self.lard/3)
        self.wazil,self.lard = 550, 800
        self.im = self.im.resize((self.wazil,self.lard), PIL.Image.ANTIALIAS)
        self.canvas.config(scrollregion=(0,0,self.wazil,self.lard))
        self.tk_im = ImageTk.PhotoImage(self.im)
        self.canvas.create_image(300, 0, anchor="nw", image=self.tk_im, )   
        #self.b = Button(root, text='开始', font=('KaiTi', 16, 'bold'), bd=2, width=8, height=5, command=func(self.url))
        #self.b.pack()
        #self.b.place(x=900, y=100, anchor=NW)

    def on_button_press(self, event):
        # save mouse drag start position
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)

        # create rectangle if not yet exist
        if not self.rect:
            self.rect = self.canvas.create_rectangle(self.x, self.y, 1, 1, outline='red')

    def on_move_press(self, event):
        curX = self.canvas.canvasx(event.x)
        curY = self.canvas.canvasy(event.y)

        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if event.x > 0.9*w:
            self.canvas.xview_scroll(1, 'units') 
        elif event.x < 0.1*w:
            self.canvas.xview_scroll(-1, 'units')
        if event.y > 0.9*h:
            self.canvas.yview_scroll(1, 'units') 
        elif event.y < 0.1*h:
            self.canvas.yview_scroll(-1, 'units')

        # expand rectangle as you drag the mouse
        self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)
        self.pos = [self.start_x, self.start_y, curX, curY]
        
    def on_button_release(self, event):
        if self.pos != [] and len(pos) < 3:
            pos.append(self.pos)
        if self.pos != []:
            self.canvas.create_rectangle(self.pos[0], self.pos[1], self.pos[2], self.pos[3], outline='red', )
        pass    
    
    def _next(self):
        if self.num < all_num - 1:
            pos.clear()
            self.pos=[]
            self.rect = None
            self.start_x = None
            self.start_y = None
            self.num += 1
            self.url = img_path[self.num]
            self.im = PIL.Image.open(self.url)
            self.wazil,self.lard = self.im.size
            self.wazil,self.lard = 550, 800
            self.im = self.im.resize((self.wazil,self.lard), PIL.Image.ANTIALIAS)
            self.canvas.config(scrollregion=(0,0,self.wazil,self.lard))
            self.tk_im = ImageTk.PhotoImage(self.im)
            self.canvas.create_image(300,0,anchor="nw",image=self.tk_im)
    
    def _previous(self):
        if self.num > 0:
            pos.clear()
            self.pos=[]
            self.rect = None
            self.start_x = None
            self.start_y = None
            self.num -= 1
            self.url = img_path[self.num]
            self.im = PIL.Image.open(self.url)
            self.wazil,self.lard = self.im.size
            self.wazil,self.lard = 550, 800
            self.im = self.im.resize((self.wazil,self.lard), PIL.Image.ANTIALIAS)
            self.canvas.config(scrollregion=(0,0,self.wazil,self.lard))
            self.tk_im = ImageTk.PhotoImage(self.im)
            self.canvas.create_image(300,0,anchor="nw",image=self.tk_im)
    
    def _delete(self):
        pos.clear()
        self.pos=[]
        self.rect = None
        self.start_x = None
        self.start_y = None
        self.url = img_path[self.num]
        self.im = PIL.Image.open(self.url)
        self.wazil,self.lard = self.im.size
        self.wazil,self.lard = 550, 800
        self.im = self.im.resize((self.wazil,self.lard), PIL.Image.ANTIALIAS)
        self.canvas.config(scrollregion=(0,0,self.wazil,self.lard))
        self.tk_im = ImageTk.PhotoImage(self.im)
        self.canvas.create_image(300,0,anchor="nw",image=self.tk_im)
        
        
    
    
def func(url):
    for i in range(len(pos)):
        if i != 0:
            im = PIL.Image.open(url)
            w, h = im.size
            im = np.array(im)
            im = im[int(pos[i][1]*h/800) : int(pos[i][3]*h/800), int((pos[i][0]-300)*w/550) : int((pos[i][2] - 300)*w/550)]
            im = PIL.Image.fromarray(im) 
            img_url = "./" + str(i) + "_img.jpg"
            im.save(img_url)
    
            client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
            i = open(img_url,'rb')
            img = i.read()
            i.close()
            message = client.basicGeneral(img)
            print(message)
            mess = []

            if message.get('words_result_num') == 0:
                mess.append(' ')
                print(' ')
            else:
                for me in message.get('words_result'):
                    print(me.get('words'))
                    mess.append(me.get('words'))
            result.append(mess)
            os.remove(img_url)
       
        
        if i == 0:
            im = PIL.Image.open(url)
            w, h = im.size
            im = np.array(im)
            im = im[int(pos[i][1]*h/800) : int(pos[i][3]*h/800), int((pos[i][0]-300)*w/550) : int((pos[i][2] - 300)*w/550)]
            im = PIL.Image.fromarray(im) 
            img_url = "./0_result.jpg"
            im.save(img_url)
            chart = []
    
            client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
            
            imageOCR = ImageTableOCR(img_url, 120, 5)
            frag_url, frag_w, frag_h = imageOCR.OCR(img_dir)
            
            for k in range(int(frag_w * frag_h)):
                i = open(frag_url[k],'rb')
                img = i.read()
                i.close()
                message = client.basicGeneral(img)
                os.remove(frag_url[k])
                if message.get('words_result_num') == 0:
                    chart.append(' ')
                    print(' ')
                else:
                    for me in message.get('words_result'):
                        print(me.get('words'))
                        chart.append(me.get('words'))
            
            result.append(chart)
    
    
    for list_num in range(len(result)):
        if list_num == 0:            
            workbook=xlwt.Workbook(encoding='utf-8')  
            booksheet=workbook.add_sheet('Sheet 1', cell_overwrite_ok=True) 
            for col in range(frag_w):
                for row in range(frag_h):
                    booksheet.write(row, col, result[0][frag_h * col + row])  
            workbook.save(img_dir + url[2:-4] + '.xls')  
        
        if list_num != 0:
            with open(img_dir + url[2:-4] + '.txt',"a+") as f:
                for mem in result[list_num]:
                    f.write(mem)
                    f.write("\n")

 


class ImageTableOCR(object):

    # 初始化
    def __init__(self, ImagePath , Threshold, Threshold_small):
        # 读取图片
        self.image = cv2.imread(ImagePath, 1)
        # 把图片转换为灰度模式
        self.gray = cv2.cvtColor(self.image, cv2.COLOR_BGR2GRAY)
        self.Threshold = Threshold
        self.frag_w = 0
        self.frag_h = 0
        self.Threshold_small = Threshold_small

    # 横向直线检测
    def HorizontalLineDetect(self):

        # 图像二值化
        ret, thresh1 = cv2.threshold(self.gray, 240, 255, cv2.THRESH_BINARY)
        # 进行两次中值滤波
        blur = cv2.medianBlur(thresh1, 3)  # 模板大小3*3
        blur = cv2.medianBlur(blur, 3)  # 模板大小3*3

        h, w = self.gray.shape

        # 横向直线列表
        horizontal_lines = []
        for i in range(h - 1):
            # 找到两条记录的分隔线段，以相邻两行的平均像素差大于120为标准
            if abs(np.mean(blur[i, :]) - np.mean(blur[i + 1, :])) > self.Threshold:
                # 在图像上绘制线段
                horizontal_lines.append([0, i, w, i])
                #cv2.line(self.image, (0, i), (w, i), (0, 255, 0), 1)

        horizontal_lines = horizontal_lines[1:]
        self.frag_h = int((len(horizontal_lines) + 2) / 2)

        return horizontal_lines

    #  纵向直线检测
    def VerticalLineDetect(self):
        # 图像二值化
        ret, thresh1 = cv2.threshold(self.gray, 240, 255, cv2.THRESH_BINARY)
        # 进行两次中值滤波
        blur = cv2.medianBlur(thresh1, 3)  # 模板大小3*3
        blur = cv2.medianBlur(blur, 3)  # 模板大小3*3

        h, w = self.gray.shape

        # 横向直线列表
        vertical_lines = []
        for i in range(w - 1):
            # 找到两条记录的分隔线段，以相邻两行的平均像素差大于120为标准
            if abs(np.mean(blur[:,i]) - np.mean(blur[:, i+1])) > self.Threshold:
                # 在图像上绘制线段
                vertical_lines.append([i, 0, i, h])
                #cv2.line(self.image, (i, 0), (i, h), (0, 255, 0), 1)

        vertical_lines = vertical_lines[1:]
        #print(vertical_lines)
        self.frag_w = int((len(vertical_lines) + 2) / 2)
        return vertical_lines

    # 顶点检测
    def VertexDetect(self):
        vertical_lines = self.VerticalLineDetect()
        horizontal_lines = self.HorizontalLineDetect()

        # 顶点列表
        vertex = []
        for v_line in vertical_lines:
            for h_line in horizontal_lines:
                vertex.append((v_line[0], h_line[1]))

        #print(vertex)

        # 绘制顶点
        for point in vertex:
            cv2.circle(self.image, point, 1, (255, 0, 0), 2)

        return vertex

    # 寻找单元格区域
    def CellDetect(self):
        vertical_lines = self.VerticalLineDetect()
        horizontal_lines = self.HorizontalLineDetect()
        #self.VertexDetect()

        # 顶点列表
        rects = []
        for i in range(0, len(vertical_lines) - 1, 2):
            for j in range(len(horizontal_lines) - 1):
                if horizontal_lines[j + 1][1] - horizontal_lines[j][1] > self.Threshold_small:
                    rects.append((vertical_lines[i][0], horizontal_lines[j][1], 
                                  vertical_lines[i + 1][0], horizontal_lines[j + 1][1]))

        # print(rects)
        return rects

    # 识别单元格中的文字
    def OCR(self, img_dir):
        rects = self.CellDetect()
        thresh = self.gray
        #cv2.imwrite("./nwe.jpg", thresh)
        frag_url = []
        frag_w, frag_h = self.frag_w - 1, self.frag_h - 1

        for j in range(frag_w):
            for m in range(frag_h):
                rect1 = rects[j * frag_h + m]
                DetectImage = thresh[rect1[1] + 1:rect1[3] + 1, rect1[0] + 1:rect1[2]]
                cv2.imwrite(img_dir + str(j + 1) + "_" + str(m + 1) + ".jpg", DetectImage)
                frag_url.append(img_dir + str(j + 1) + "_" + str(m + 1) + ".jpg")
        
        return frag_url, frag_w, frag_h
             

    # 显示图像
    def ShowImage(self):
        cv2.imshow('AI', self.image)
        cv2.waitKey(0)
        #cv2.imwrite('E://Horizontal.png', self.image)
        cv2.destroyAllWindows()




    
if __name__ == "__main__":
    APP_ID = '16727161'
    API_KEY = '6w7HWTctY7CaDWICoITA2g4b'
    SECRET_KEY = 'b79Tu2kYUopsZZn1gDncmRd5mCBY1nYC'
    img_path = glob.glob('./*.jpg')
    img_dir="./"
    result_path = "./"
    
    all_num = len(img_path)
    root = Tk()
    pos = []
    result = []
    root.geometry("1200x1200")
    app = MainApp(root)
    app.pack()
    

    b1 = Button(root, text='开始', font=('KaiTi', 16, 'bold'), bd=2, width=8, height=5, command = lambda : func(app.url),)
    b1.pack()
    b1.place(x=900, y=100, anchor=NW)
    b2 = Button(root, text='下一张', font=('KaiTi', 16, 'bold'), bd=2, width=8, height=5, command = lambda :app._next())
    b2.pack()
    b2.place(x=900, y=300, anchor=NW)
    b3 = Button(root, text='上一张', font=('KaiTi', 16, 'bold'), bd=2, width=8, height=5, command = lambda :app._previous())
    b3.pack()
    b3.place(x=150, y=300, anchor=NW)
    b4 = Button(root, text='清空', font=('KaiTi', 16, 'bold'), bd=2, width=8, height=5, command = lambda :app._delete())
    b4.pack()
    b4.place(x=900, y=500, anchor=NW)
    
    root.mainloop()
    
    
    
