import win32com.client as wc


from tkinter import *
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import messagebox
root=Tk();
root.geometry('600x600')
root.title("Deepak-Note")
textpad=scrolledtext.ScrolledText(root,width=100,height=100)
def openc():
    file=filedialog.askopenfile(parent=root,mode='r',title="select a file")
    if(file!=None):
        contents=file.read()
        textpad.insert('1.0',contents)
        file.close()
def savec():
    file=filedialog.asksaveasfile(mode='w')
    if(file!=None):
        data=textpad.get('1.0',END+'-1c')
        file.write(data)
        
        file.close()
def speak():
    speak=wc.Dispatch("SAPI.SpVoice")
    while(1):
        a=textpad.get('1.0',END+'-1c')
        speak.Speak(a)
        break
def exitt():
    if messagebox.askokcancel("Quit","Do you want to Exit?"):
        root.destroy()
def tt():
    poot=Tk()
    poot.minsize(width=500,height=500)
    poot.maxsize(width=500,height=500)
    msgg=Message(poot,text="this is a simple text editor \n created by deepak tiwari")
    msgg.config(bg='pink',font=('times',24,'italic'))
    msgg.pack()
    poot.mainloop()
menu=Menu(root)
root.config(menu=menu)
filemenu=Menu(menu)
filemenu1=Menu(menu)
menu.add_cascade(label="File",menu=filemenu)
menu.add_cascade(label="Edit")
menu.add_cascade(label="Format")
menu.add_cascade(label="About",command=tt)
menu.add_cascade(label="speak",command=speak)
filemenu.add_command(label="New")
filemenu.add_command(label="Open",command=openc)
filemenu.add_command(label="Save",command=savec)
filemenu.add_separator()
filemenu.add_command(label="Exit",command=exitt)
textpad.pack(fill=Y)

root.mainloop()
