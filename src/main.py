#!/usr/bin/python3
#-*- coding:utf-8 -*-
'''
Created on 2016年7月19日

@author: mikewolfli
'''
from tkinter import *
from tkinter import simpledialog
from tkinter import font
from tkinter import scrolledtext
from tkinter import messagebox
from tkinter import filedialog
import tkinter.ttk as ttk
from ldap3 import Server, Connection, SIMPLE, SYNC, ALL, SASL, NTLM
from dataset import *
import logging 
import datetime

def center(toplevel):
    toplevel.update_idletasks()
    w = toplevel.winfo_screenwidth()
    h = toplevel.winfo_screenheight()
    size = tuple(int(_) for _ in toplevel.wm_geometry().split('+')[0].split('x'))
    x = w/2 - size[0]/2
    y = h/2 - size[1]/2
    toplevel.wm_geometry("%dx%d+%d+%d" % (size + (x, y)))

class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        self.formatter = logging.Formatter('%(asctime)s-%(levelname)s : %(message)s')
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')          
            self.text.insert(END, msg+"\n")
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)# Scroll to the bottom
        
login_info ={'uid':'','pwd':'','status':False,'perm':'0000'}
        
class LoginForm(Toplevel):
    def __init__(self, parent, title=None):
        Toplevel.__init__(self, parent)

        self.withdraw()
        if parent.winfo_viewable():
            self.transient(parent)

        if title:
            self.title(title)

        self.parent = parent
        #self.grab_set()
        body = Frame(self)
        self.initial_focus = self.body(body)
        body.pack(padx=5, pady=5)

        self.buttonbox()

        if not self.initial_focus:
            self.initial_focus = self

        self.protocol("WM_DELETE_WINDOW", self.cancel)

        if self.parent is not None:
            center(self)
            
            #self.geometry("+%d+%d" % (parent.winfo_rootx()+50,
            #                          parent.winfo_rooty()+50))

        self.deiconify() # become visible now

        self.initial_focus.focus_set()

        # wait for window to appear on screen before calling grab_set
        self.wait_visibility()
        
        #self.grab_set()
        
        self.wait_window(self) 
            
    def body(self, master):
        self.label1 = Label(master,text='用户名')
        self.label1.grid(row=0, column=0, sticky=W)
        self.uid_entry= Entry(master)
        self.uid_entry.grid(row=0, column=1, columnspan=2,  sticky=EW)       
        self.label2 = Label(master, text='密码')
        self.label2.grid(row=1, column=0, sticky=W)
        self.pwd_entry = Entry(master, show='*')
        self.pwd_entry.grid(row=1, column=1, columnspan=2, sticky=EW)
        self.msg_str=StringVar()
        self.label_message = Label(master, textvariable=self.msg_str).grid(row=2, column=0, columnspan=3, sticky=W)
        return self.uid_entry

    def buttonbox(self):
        box = Frame(self)

        w = Button(box, text="登陆", width=10, command=self.ok, default=ACTIVE)
        w.pack( side=LEFT, padx=5, pady=5)
        w = Button(box, text="取消", width=10, command=self.cancel)
        w.pack(side=LEFT, padx=5, pady=5)

        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)

        box.pack()
        
    def validate(self):  
        login_info['uid'] = self.uid_entry.get()
        login_info['pwd'] = self.pwd_entry.get()
        
        if len(login_info['uid'].strip())==0:
            self.msg_str.set('用户名不能为空')
            self.initial_focus=self.uid_entry
            self.initial_focus.focus_set()
            return 0
            
        s = Server('tkeasia.com', get_info=ALL)
        c = Connection(s, user="tkeasia\\"+login_info['uid'], password=login_info['pwd'], authentication=NTLM)
        ret = c.bind()
        if ret:
            login_info['status'] = True
            self.get_permission()
            self.log_login()
            return 1
        else:
            self.msg_str.set('登陆失败！')
            messagebox.showerror('错误', c.last_error)
            return 0
        
    def destroy(self):
        self.initial_focus = None
        Toplevel.destroy(self)    
        
    def cancel(self, event=None):
        if self.parent is not None:
            self.parent.focus_set()
        self.destroy()
        
        if pg_db.get_conn():
            pg_db.close()
            
        if mbom_db.get_conn():          
            mbom_db.close()
            
        sys.exit()
        
    def ok(self, event=None):
        if not self.validate():
            self.initial_focus.focus_set() # put focus back
            return
        self.withdraw()
        self.update_idletasks()        
        self.destroy()
        
    def log_login(self):
        if not mbom_db.get_conn():
            mbom_db.connect()
              
        login_record = login_log.select().where((login_log.employee==login_info['uid'])&(login_log.log_status==True))
        if len(login_record) >0:   
            log_loger = login_log.update(log_status=False).where((login_log.employee==login_info['uid'])&(login_log.log_status==True))
            log_loger.execute()
        
        
        log_loger = login_log.insert(employee=login_info['uid'], log_status=True, login_time=datetime.datetime.now())
        log_loger.execute()
    
    def get_permission(self):
        try:
            perm = op_permission.get(op_permission.employee==login_info['uid'])
            login_info['perm']= perm.perm_code
        except op_permission.DoesNotExist:
            pass

class mainframe(Frame):
    def __init__(self,master=None):
        Frame.__init__(self, master)
        self.pack()
        
    def createWidgets(self):
        pass

class Application():
    def __init__(self, root):      
        main_frame = mainframe(root)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.grid(row=0, column=0, sticky=NSEW)
        LoginForm(main_frame,'用户登陆')
        root.protocol("WM_DELETE_WINDOW", self.quit_func)

    def quit_func(self):          
        if pg_db.get_conn():
            pg_db.close()
            
        if mbom_db.get_conn() and login_info['status']:
            log_loger = login_log.update(logout_time = datetime.datetime.now(), log_status=False).where((login_log.employee==login_info['uid'])&(login_log.log_status==True))
            log_loger.execute()            
            mbom_db.close()
        root.destroy()
                
if __name__ == '__main__':   
    root=Tk() 
    #root.resizable(0, 0)
    
    #root.attributes("-zoomed", 1) # this line removes the window managers bar       
        
    try:                                   # Automatic zoom if possible
        root.wm_state("zoomed")
    except TclError:                    # Manual zoom
        # get the screen dimensions
        width = root.winfo_screenwidth()
        height = root.winfo_screenheight()

        # borm: width x height + x_offset + y_offset
        geom_string = "%dx%d+0+0" % (width, height)
        root.wm_geometry(geom_string)
        
    root.title('物料转换')
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(size=10)  
    root.option_add("*Font", default_font)
    Application(root)
    root.geometry('800x600')
    root.mainloop()