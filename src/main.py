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
import pandas as pd
from pandastable import Table, TableModel
from ldap3 import Server, Connection, SIMPLE, SYNC, ALL, SASL, NTLM
from dataset import *
import logging 
import datetime
from openpyxl import *
from treelib import Tree, Node
from decimal import Decimal
from numpy.core.defchararray import isdecimal

NAME = 'EDS非标物料处理'
PUBLISH_KEY=' A ' #R - release , B - Beta , A- Alpha
VERSION = '0.1.0'

logger = logging.getLogger()

def value2key(dic, value):
    if not isinstance(dic, dict):
        return None
    
    for key, val in dic.items():
        if val == value:
            return key
    
    return None

def date2str(dt_s):
    if not isinstance(dt_s, datetime.datetime):
        return None
    else:
        return dt_s.strftime("%Y-%m-%d") 

def datetime2str(dt_s):
    if not isinstance(dt_s, datetime.datetime):
        return None
    else:
        return dt_s.strftime("%Y-%m-%d %H:%M:%S") 
    
def str2date(dt_s):
    if dt_s is None or (len(dt_s)==0 and isinstance(dt_s, str)):
        return None
    else:
        return datetime.datetime.strptime(dt_s , '%Y-%m-%d')
    
def str2datetime(dt_s):
    if dt_s is None or (len(dt_s)==0 and isinstance(dt_s, str)) :
        return None
    else:
        return datetime.datetime.strptime(dt_s, "%Y-%m-%d %H:%M:%S") 

def none2str(val):
    if not val:
        return ''
    else:
        return val.rstrip()

def center(toplevel):
    toplevel.update_idletasks()
    w = toplevel.winfo_screenwidth()
    h = toplevel.winfo_screenheight()
    size = tuple(int(_) for _ in toplevel.geometry().split('+')[0].split('x'))
    x = w/2 - size[0]/2
    y = h/2 - size[1]/2
    toplevel.geometry("%dx%d+%d+%d" % (size + (x, y)))

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
        
login_info ={'uid':'','pwd':'','status':False,'perm':'0000','plant':'2101'}
        
class LoginForm(Toplevel):
    def __init__(self, parent, title=None):
        Toplevel.__init__(self, parent)

        self.withdraw()
        if parent.winfo_viewable():
            self.transient(parent)

        if title:
            self.title(title)

        self.parent = parent
        body = Frame(self)
        self.initial_focus = self.body(body)
        body.pack(padx=5, pady=5)

        self.buttonbox()

        if not self.initial_focus:
            self.initial_focus = self

        self.protocol("WM_DELETE_WINDOW", self.cancel)

        center(self)
        #if self.parent is not None:         
            #self.geometry("+%d+%d" % (parent.winfo_rootx()+50,
            #                         parent.winfo_rooty()+50))

        self.deiconify() # become visible now

        self.initial_focus.focus_set()

        # wait for window to appear on screen before calling grab_set
        self.wait_visibility()
        
        self.grab_set()
        
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
            
            logger.info(login_info['uid']+"登陆系统...")
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

mat_heads = ['位置号','物料号','中文名称','英文名称','图号','数量','单位','材料','重量','备注','非标判断']
mat_keys = ['st_no','mat_no','mat_name_cn','mat_name_en','drawing_no','qty','mat_unit','mat_material','part_weight','comments','is_nonstd']

mat_cols = ['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11']

class mainframe(Frame):
    '''
    mat_list = {1:{'位置号':value,'物料号':value, ....,'标判断':value},.....,item:{......}}
    mat_tree : 物料BOM的树形结构， 取t_list的key,如下:
    0
    ├── 1
    │   └── 3
    └── 2
    '''
    mat_list = {}
    struct_code=''
    bom_tree = Tree()
    bom_tree.create_node(0,0)
    mat_pos = 0
    mat_tops={}
    def __init__(self,master=None):
        Frame.__init__(self, master)
        self.pack()
        
        self.createWidgets()
        
    def createWidgets(self):
        self.find_mode = StringVar()
        self.find_combo = ttk.Combobox(self,textvariable = self.find_mode)
        self.find_combo['values'] = ('列出物料BOM结构','查找物料的上层','查找物料关联项目','查找项目关联物料')
        self.find_combo.current(0)
        self.find_combo.grid(row =0,column=0, columnspan=2,sticky=EW)
        
        self.find_var = StringVar()
        self.find_text = Entry(self, textvariable=self.find_var)
        self.find_text.grid(row=1, column=0, columnspan=2, sticky=EW)
                
        st_body = Frame(self)
        st_body.grid(row=0, column=2,sticky=NSEW)
        
        self.import_button = Button(st_body, text='文件读取')
        self.import_button.pack(side='left')
        self.import_button['command'] = self.excel_import
        
        self.check_sap = Button(st_body, text='非标物料系统比对')
        self.check_sap.pack(side='left')
        
        self.generate_nstd_list = Button(st_body, text='生成非标物料申请表')
        self.generate_nstd_list.pack(side='left')
               
        ie_body = Frame(self)
        ie_body.grid(row=1, column=2, columnspan=10, sticky=NSEW)
        
        self.import_bom_List = Button(ie_body, text='生成BOM导入表')
        self.import_bom_List.pack(side='left')
        
        self.ntbook = ttk.Notebook(self)        
        self.ntbook.rowconfigure(0, weight=1)
        self.ntbook.columnconfigure(0, weight=1)
        
        list_pane = Frame(self)
        model = TableModel(rows=0, columns=0)
        for col in mat_heads:
            model.addColumn(col)
        model.addRow(1)
        
        self.mat_table = Table(list_pane, model, editable=False)
        self.mat_table.show()
                 
        tree_pane = Frame(self)
        self.mat_tree = ttk.Treeview(tree_pane, show='headings', columns=mat_cols,selectmode='extended')
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))  
        self.mat_tree.heading('#0', text='')
        for col in mat_cols:
            i = mat_cols.index(col)
            self.mat_tree.heading(col,text=mat_heads[i])
        
        #('位置号','物料号','中文名称','英文名称','图号','数量','单位','材料','重量','备注')
        self.mat_tree.column('#0', width=20)
        self.mat_tree.column('col1', width=80, anchor='w')
        self.mat_tree.column('col2', width=100, anchor='w')
        self.mat_tree.column('col3', width=150, anchor='w')
        self.mat_tree.column('col4', width=150, anchor='w')
        self.mat_tree.column('col5', width=100, anchor='w')
        self.mat_tree.column('col6', width=100, anchor='w')
        self.mat_tree.column('col7', width=100, anchor='w')
        self.mat_tree.column('col8', width=150, anchor='w')
        self.mat_tree.column('col9', width=100, anchor='w')
        self.mat_tree.column('col10', width=300, anchor='w')      
               
        ysb = ttk.Scrollbar(tree_pane, orient='vertical', command=self.mat_tree.yview)
        xsb = ttk.Scrollbar(tree_pane, orient='horizontal', command=self.mat_tree.xview)        
        ysb.grid(row=0, column=2, rowspan=2, sticky='ns')
        xsb.grid(row=2, column=0, columnspan=2, sticky='ew')  
        
        self.mat_tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.mat_tree.grid(row=0, column=0, rowspan=2, columnspan=2, sticky='nsew')
        tree_pane.rowconfigure(1, weight=1)
        tree_pane.columnconfigure(1, weight =1)
        
        self.ntbook.add(list_pane, text='BOM清单', sticky=NSEW)
        self.ntbook.add(tree_pane, text='BOM树形结构', sticky=NSEW) 
        
        log_pane = Frame(self)
        
        self.log_label=Label(log_pane)
        self.log_label["text"]="操作记录"
        self.log_label.grid(row=0,column=0, sticky=W)
        
        self.log_text=scrolledtext.ScrolledText(log_pane, state='disabled')
        self.log_text.config(font=('TkFixedFont', 10, 'normal'))
        self.log_text.grid(row=1, column=0, columnspan=2, sticky=EW)
        log_pane.rowconfigure(1,weight=1)
        log_pane.columnconfigure(1, weight=1)
        
        self.ntbook.grid(row=2, column=0, rowspan=6, columnspan=6, sticky=NSEW)
        log_pane.grid(row=8, column=0,columnspan=6, sticky=NSEW)
              
        # Create textLogger
        text_handler = TextHandler(self.log_text)        
        # Add the handler to logger
        
        logger.addHandler(text_handler)
        logger.setLevel(logging.INFO) 
        
        self.rowconfigure(8, weight=1)
        self.columnconfigure(5, weight=1) 
        
    def excel_import(self):
        file_list = filedialog.askopenfilenames(title="导入文件", filetypes=[('excel file','.xlsx'),('excel file','.xlsm')])
        if not file_list:
            return

        self.mat_list = {}
        self.struct_code=''
        self.mat_pos = 0
        self.mat_tops = {}
        for node in self.bom_tree.children(0):
            self.bom_tree.remove_node(node.identifier)

        for file in file_list:
            logger.info("正在读取文件:"+file+",转换保存物料信息,同时构建数据Model")
            c=self.read_excel_files(file)
            logger.info("文件:"+file+"读取完成, 共计处理 "+str(c)+" 个物料。")
            
        df = pd.DataFrame(self.mat_list,index=mat_heads, columns=[ i for i in range(1, self.mat_pos+1)])
        model = TableModel(dataframe=df.T)
        self.mat_table.updateModel(model)
        self.mat_table.redraw()
        
    def save_mat_info(self,method=False,**para):
        try:
            mat_info.get(mat_info.mat_no == para['mat_no'])
            if method:
                q = mat_info.update(mat_name_en=para['mat_name_en'], mat_name_cn=para['mat_name_cn'], drawing_no=para['drawing_no'],mat_material=para['mat_material'],mat_unit=para['mat_unit'],\
                                mat_material_en=para['mat_material_en'],part_weight=para['part_weight'],comments=para['comments'],modify_by=login_info['uid'],modify_on=datetime.datetime.now()).where(mat_info.mat_no==para['mat_no']) 
                return q.execute()                          
        except mat_info.DoesNotExist:
            q = mat_info.insert(mat_no=para['mat_no'], mat_name_en=para['mat_name_en'], mat_name_cn=para['mat_name_cn'], drawing_no=para['drawing_no'],mat_material=para['mat_material'],mat_unit=para['mat_unit'],\
                                mat_material_en=para['mat_material_en'],part_weight=para['part_weight'],comments=para['comments'],modify_by=login_info['uid'],modify_on=datetime.datetime.now())
            return q.execute()
        
        return 0
        
    def build_tree_struct(self):
        pass
 
    def read_excel_files(self, file):
        '''
                返回值：
                -2: 读取EXCEL失败
                -1 : 头物料位置为空
                0： 头物料的版本已经存在
                1： 
       
        '''
        wb = load_workbook(file, read_only=True, data_only=True)
        sheetnames=wb.get_sheet_names()
        
        if len(sheetnames)==0:
            return -2
        
        counter=0
        for i in range(len(sheetnames)): 
            for j in range(1,19):                         
                mat_line = {}
                mat_top_line={}
                ws = wb.get_sheet_by_name(sheetnames[i]) 
            
                mat_line[mat_heads[0]]=none2str(ws.cell(row=2*j+1,column=2))
                mat_line[mat_heads[1]]=none2str(ws.cell(row=2*j+1,column=5))
                if len(mat_line[mat_heads[1]])==0:
                    break
                           
                mat_line[mat_heads[2]]=none2str(ws.cell(row=2*j+1,column=7))
                mat_line[mat_heads[3]]=none2str(ws.cell(row=2*j+2,column=7))
                mat_line[mat_heads[4]] = none2str(ws.cell(row=2*j+1, column=6))
                
                qty = none2str(ws.cell(row=2*j+1,column=3))
                if len(qty)==0 or Decimal(qty)==0: 
                    continue
                
                self.mat_pos+=1
                counter+=1
                
                mat_line[mat_heads[5]] = Decimal(qty)
                mat_line[mat_heads[6]]=none2str(ws.cell(row=2*j+1,column=4))
                mat_line[mat_heads[7]] = none2str(ws.cell(row=2*j+1,column=9))
                material_en = none2str(ws.cell(row=2*j+2, column=9))
                
                weight = none2str(ws.cell(row=2*j+1, column=10))
                if len(weight)==0:
                    mat_line[mat_heads[8]]=0
                else:
                    mat_line[mat_heads[8]]=Decimal(weight)
                    
                mat_line[mat_heads[9]]=none2str(ws.cell(row=2*j+1, column=11))
                
                #保存物料基本信息
                if self.save_mat_info(mat_no=mat_line[mat_heads[1]], mat_name_en=mat_line[mat_heads[3]], mat_name_cn=mat_line[mat_heads[2]], drawing_no=mat_line[mat_heads[4]],mat_material=mat_line[mat_heads[7]],mat_unit=mat_line[mat_heads[6]],\
                                mat_material_en=material_en,part_weight=mat_line[mat_heads[8]],comments=mat_line[mat_heads[9]])==0:
                    logger.info(mat_line[mat_heads[1]]+'数据库中已经存在,故没有保存')
                else:
                    logger.info(mat_line[mat_heads[1]]+'保存成功。')
                                
                if i==0 and j==1:
                    mat_top_line['revision'] = none2str(ws.cell(row=43, column=8))
                    mat_top_line['struct_code']=none2str(ws.cell(row=39,column=12))
                    
                    self.mat_tops[mat_line[mat_heads[1]]]=mat_top_line
                
                self.mat_list[self.mat_pos] = mat_line
                
        return counter
                
                                      
    def bom_id_generator(self):
        try:
            bom_res = id_generator.get(id_generator.id == 1)
        except id_generator.DoesNotExist:
            return None
        
        pre_char = bom_res.pre_character
        fol_char = bom_res.fol_character
        c_len = bom_res.id_length
        cur_id = bom_res.id_length
        step = bom_res.step        
        new_id=str(cur_id+step)
        #前缀+前侧补零后长度为c_len+后缀, 组成新的BOM id               
        id_char = pre_char+new_id.zfill(c_len)+fol_char
        
        q=id_generator.update(current=cur_id+step).where(id_generator.id==1)
        q.execute()
        
        return id_char
                              
class Application():
    def __init__(self, root):      
        main_frame = mainframe(root)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.grid(row=0, column=0, sticky=NSEW)
        LoginForm(root,'用户登陆')
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
        
    root.title(NAME+PUBLISH_KEY+VERSION)
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(size=10)  
    root.option_add("*Font", default_font)
    Application(root)
    root.geometry('800x600')
    root.mainloop()