import tkinter
import tkinter.filedialog
import CheckAttendance
import tkinter.messagebox #弹窗库
#version: 3.0
#date: 20190801

win = tkinter.Tk()
 
win.title('FC-考勤制表工具')
win.geometry('500x400')
win.resizable()

class GuiAction:
	filePath =''
	
	#获取文件的位置
	def readFile(self):
		global filePath 
		filePath = tkinter.filedialog.askopenfilename()
		
		if filePath != '':
			lb.config(text = "选择的文件是："+filePath);
			savebtn.configure(state='active') 
		else:
			lb.config(text = "FC，你没有选择任何文件");
		
	#运行并保存文件至选择的路径
	def saveFile(self):
		savePath = tkinter.filedialog.asksaveasfilename(title=u'保存文件', filetypes=[("XLSX", ".xlsx")])
		CheckAttendance.Attendance().runCheckAttendance(filePath, savePath)
		self.messageWindow()
	
	#自定义消息弹窗	
	def messageWindow(self):
		messageWin = tkinter.Toplevel()
		messageWin.geometry('300x200')
		messageWin.title('提示')
		message = "本月考勤表已制作完成\r\n\n\n\nFC，我爱你"
		tkinter.Label(messageWin, text=message).pack()
		tkinter.Button(messageWin, text='我也是', command=win.quit).pack()
	

#顶上空一行
lb = tkinter.Label(win,text = '')
lb.pack()


#文件选择按钮
readbtn = tkinter.Button(win,text="选择",width=10,height=4,command=GuiAction().readFile)
readbtn.pack()



lb1 = tkinter.Label(win,text = '')
lb2 = tkinter.Label(win,text = '')
lb3 = tkinter.Label(win,text = '')
lb4 = tkinter.Label(win,text = '')
lb5 = tkinter.Label(win,text = '')
lb1.pack()
lb2.pack()
lb3.pack()
lb4.pack()
lb5.pack()


#文件保存按钮
savebtn = tkinter.Button(win,text="保存", width=10,height=4,command=GuiAction().saveFile)
savebtn.configure(state='disabled') 
savebtn.pack()


# 进入消息循环
win.mainloop()



