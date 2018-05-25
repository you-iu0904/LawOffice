import Attorney as a
import tkinter.filedialog
import tkinter as tk
import sys
import os
window = tk.Tk()
window.title('xxx律师所')
window.geometry('500x300')
window.maxsize(500, 300)
window.minsize(500, 300)

var_userFile = tk.StringVar()
var_billsFile = tk.StringVar()
var_stageFile = tk.StringVar()


def chooseuserfile(event):
    userfile = tk.filedialog.askopenfilename()
    var_userFile.set(userfile)


def choosebillsfile(event):
    billsfile = tk.filedialog.askopenfilename()
    var_billsFile.set(billsfile)


def choosestagefile(event):
    stagefile = tk.filedialog.askopenfilename()
    var_stageFile.set(stagefile)


def confirm():
    if var_userFile.get() == '':
        tk.messagebox.showinfo(title='提示', message='用户保存路径不能为空')
    elif var_billsFile.get() == '':
        tk.messagebox.showinfo(title='提示', message='单据保存路径不能为空')
    elif var_stageFile.get() == '':
        tk.messagebox.showinfo(title='提示', message='Stage保存路径不能为空')
    elif var_userFile.get().endswith('.txt') != True or var_billsFile.get().endswith('.txt') != True or var_stageFile.get().endswith('.txt') != True:
        tk.messagebox.showinfo(title='提示', message='文件格式必须为.txt格式')
    elif os.path.exists(var_userFile.get())==False:
        tk.messagebox.showinfo(title='提示',message='用户文件不存在,请重新选择')
    elif os.path.exists(var_billsFile.get())==False:
        tk.messagebox.showinfo(title='提示', message='单据文件不存在,请重新选择')
    elif os.path.exists(var_stageFile.get())==False:
        tk.messagebox.showinfo(title='提示', message='Stage文件不存在,请重新选择')
    else:
        data_dict={}
        data_list = []
        data_list.append(var_userFile.get())
        data_list.append(var_billsFile.get())
        data_list.append(var_stageFile.get())
        data_dict['record']=data_list
        recordFile=open('Data/record.txt', 'w')
        recordFile.write(str(data_dict))
        recordFile.close()
        window.destroy()
        a.mian(var_userFile.get(), var_billsFile.get(), var_stageFile.get())
recordFile = open('Data/record.txt', 'r')
text=recordFile.readlines()
for dic in text:
    data=eval(dic)
    for i in data.values():
        var_userFile.set(i[0])
        var_billsFile.set(i[1])
        var_stageFile.set(i[2])
def cancel():
    sys.exit()

canvas = tk.Canvas(window, height=200, width=500)
image_file = tk.PhotoImage(file='image/welcome.jpg')
image = canvas.create_image(0, 0, anchor='nw', image=image_file)
canvas.place(x=20, y=20)

userfile_label = tk.Label(window, text='用户保存路径:').place(x=30, y=200)
billsfile_label = tk.Label(window, text='单据保存路径:').place(x=30, y=230)
stagefile_label = tk.Label(window, text='Stage保存路径:').place(x=30, y=260)

userfile_entry = tk.Entry(window, textvariable=var_userFile, width=35)

userfile_entry.bind('<Double-Button-1>', chooseuserfile)
userfile_entry.place(x=150, y=203)

billsfile_entry = tk.Entry(window, textvariable=var_billsFile, width=35)
billsfile_entry.bind('<Double-Button-1>', choosebillsfile)
billsfile_entry.place(x=150, y=233)

stagefile_entry = tk.Entry(window, textvariable=var_stageFile, width=35)
stagefile_entry.bind('<Double-Button-1>', choosestagefile)
stagefile_entry.place(x=150, y=263)

confirm = tk.Button(window, width=6, text='确定', command=confirm).place(x=420, y=202)
cancel = tk.Button(window, width=6, text='退出', command=cancel).place(x=420, y=250)

window.mainloop()
