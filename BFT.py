# Solved, # Bug-0, Not able to save file, when the window is closing. Solved using Path.txt
# Solution is to do the file saving normally, but only if the that selected path have log file. This can be achieved by using a function just before saving the file, to check True or False.

# Solved # Bug-1, when the gui is opened, a folder (with log) can be created and it reflects in the gui, but when it is deleted, the last button in the gui gets duplicated.
# Solved # Bug-2, mousescroll of left window also fired when the mouse is scrolled in the right window. Soled using Enter and Leave event.
# Solved # Bug-3, even when we select a path with folders that doesn't contain any log files, the path gets successfully added; it shows no data, but the prev button will still be visible. But, if the gui is closed in such a position, it won't load up again.
# Solved using 'folders_check' function iteratively and using file_check instead of file_real inside the function message_()
# Bug-4, when the gui is opened for a long time, it glitches or freezes for some time. Need to implement threading to solve this problem.
# Implement-1, Need to add dowwnloadable excel to treeview
# Bug-5, when the log.txt is updated when the gui is still on, the button position change gets updated, but the treeview labels gets updated only when search/reset of click any other button and come back.
# Solved # Bug-6 continuation of Bug-1, when test server 2 is selected with fewer buttons than test server 1, the old buttons didn,t get deleted.
# 1,6 solved using a introducing a list that contains all the buttons in the path, and deleting as required.
# Implement-2, Need to add search functionality to the bft buttons.
# Implement-3, Need to add highlight button feature to the bft buttons.
# Implement-4, Need to style the entire display.
# Solved # Bug-7, the treeview headers should refelect the headers in the csv file, if any change occurs. Solved using a list to append data.
# Solved # Bug-8, when trying to implement color code the treeview doesn't show up.
# Bug-9, Have to remove the infinite loop by implementing a Data Generate button.

import os
import pandas as pd
import datetime
from datetime import datetime
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter.messagebox import showinfo
from tkinter import filedialog
import threading
from openpyxl.workbook import Workbook # To create spreadsheets
from openpyxl import load_workbook # To access spreadsheets
import csv
import webbrowser
from PIL import Image,ImageTk

root =Tk()
root.geometry('1200x500')

#root.grid_rowconfigure(1, weight=1)
#root.grid_columnconfigure(0, weight=1)

main_Panel = PanedWindow(root, background="black")
root.grid_rowconfigure(0, weight=50)
root.grid_columnconfigure(0, weight=50)
main_Panel.grid(row=0,column=0, sticky="nsew") # pack(side="top", fill="both", expand=True)

# Create two frames
left_pane = Frame(main_Panel, background="grey", width=150)
right_pane = Frame(main_Panel, background="grey", width=200)
main_Panel.add(left_pane)
main_Panel.add(right_pane)

#status_bar = Label(main_Panel,text="test",bd=1,relief=SUNKEN,anchor=W)
#main_Panel.add(status_bar)

# Create canvas for the left frame.
left_canvas = Canvas(left_pane,background="grey",width=275)
left_canvas.pack(side=LEFT,fill=BOTH, expand=1) #grid(rowspan=1000,columnspan=100,sticky="nsew") # May be try with pack after adding frame
#left_pane.grid_rowconfigure(1, weight=1)
#left_pane.grid_columnconfigure(1, weight=1)

# Scrollbar for bft/canvas
btf_scroll = ttk.Scrollbar(left_pane, orient=VERTICAL, command = left_canvas.yview)
btf_scroll.pack(side=RIGHT,fill=Y)  #grid(rowspan=10,column=5)

# Configure canvas
left_canvas.configure(yscrollcommand=btf_scroll.set)
left_canvas.bind('<Configure>',lambda e: left_canvas.configure(scrollregion = left_canvas.bbox("all")))
#left_canvas.bind("<MouseWheel>", _on_mousewheel)

def _bound_to_mousewheel(event):
    left_canvas.bind_all("<MouseWheel>", _on_mousewheel)

def _unbound_to_mousewheel(event):
    left_canvas.unbind_all("<MouseWheel>")

def _on_mousewheel(event):
    left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

left_canvas.bind('<Enter>', _bound_to_mousewheel)
left_canvas.bind('<Leave>', _unbound_to_mousewheel)

# Create frame for left canvas
Inner_frame = Frame(left_canvas,background="grey", width=150)
Inner_frame.grid_columnconfigure(0, weight=1)
#Inner_frame.grid_columnconfigure(1, weight=1)
#Inner_frame.grid_columnconfigure(2, weight=1)
#Inner_frame.grid_columnconfigure(3, weight=1)
#Inner_frame.grid_columnconfigure(5, weight=1)
#Inner_frame.grid_columnconfigure(5, weight=1)
#Inner_frame.grid_columnconfigure(6, weight=1)
#Inner_frame.grid_columnconfigure(7, weight=1)

button_bft = Button(Inner_frame)

Inner_frame.bind('<Enter>',lambda e:left_canvas.bind("<MouseWheel>", _on_mousewheel))

# Adding the inner frame to a window in the canvas.
left_canvas.create_window((0,0), window=Inner_frame, anchor="nw")


label_1 = Label(Inner_frame,text="Recently Updated BFTs",font=("Times", 15))
label_1.grid(row=0,column=0,pady=10)

#label_1.bind("<Button-1>", lambda e:
#callback("https://guihelp.coolove.repl.co/"))

my_menu = Menu(root)
root.config(menu=my_menu)

# Creating menu item
help_menu = Menu(my_menu,tearoff=0)
my_menu.add_cascade(label="Help", menu = help_menu)
help_menu.add_command(label = "Get Help", command=lambda: callback("https://guihelp2.coolove.repl.co/"))

#Define a callback function
def callback(url):
   webbrowser.open_new_tab(url)

#help_menu.bind("<Button-1>", lambda e:
#callback("https://guihelp.coolove.repl.co/"))

p=''
Btn_to_del = {}
flag = False
buttons = []
selected_button = None
last_bg = None
#wb = Workbook() # Creating workbook instance

try: # Checking if path.txt exists
    file_real = open("Path.txt","r")
    file_real.close
except: # If path.txt doesn't exist, it is created.
    file_real = open("Path.txt","w")
    file_real.close

def check_folders(check_path):
    global file_final
    #print('check_folders called.')
    global p
    #print(f'p inside check_folders: {p}')
    folder_check_list = [] # Folder check list
    folder_path_check_list = [] # Folder path check list
    folder_path_log_check_list = [] # List containing log file path. 
    folders_check = os.listdir(check_path) # Folders
    #print(folders_check)
    for folder in folders_check:
        folder_check_list.append(folder)
        folder_path = os.path.join(check_path,folder)
        folder_path_check_list.append(folder_path)

    #print(folder_path_check_list)

    for path in folder_path_check_list:
        if (os.path.isfile(os.path.join(path,'log.txt'))):
            path_log = os.path.join(path,'log.txt')
            folder_path_log_check_list.append(path_log)
            #time_mod = os.stat(path_log).st_mtime
            #time_list.append(time_mod)
        else:
            folder_path_check_list.remove(path)
            folder_check_list.remove(os.path.basename(path))

    if (folder_path_log_check_list==[]):
        #print('called')
        messagebox.showerror("Path Invalid", "Please select valid path")
        drop.configure(state="disabled")
        file_final = open("Path.txt","w")
        file_final.write(p)
        file_final.close
        Delete_Button()
        #select_path()
    else:
        #print('Check_folders --> Open_path called.')
        Open_path(p)

def select_when_empty():
    global p
    global file_final
    path_mainDir = filedialog.askdirectory()
    #(path)
    path_mainDir = path_mainDir.replace("/","\\")
    p = path_mainDir
    #print(p)
    if p=='':
        messagebox.showerror("Path Invalid", "Please select valid path")
        drop.configure(state="disabled")
        select_path()
        file_final = open("Path.txt","w")
        file_final.write(p)
        file_final.close
    else:
        # Save it in another fn or variable.
        file_final = open("Path.txt","w")
        file_final.write(p)
        file_final.close
        Delete_Button()
        check_folders(p)

def message_():
    global p
    global file_check
    file_check = open("Path.txt","r")
    p = file_check.read()
    file_check.close()
    #print(f'p inside message after reading: {p}')
    if p=='':      
        file_check = open("Path.txt","r")
        if os.stat("Path.txt").st_size == 0:
            file_check.close
            #print('message_called')
            #messagebox.showinfo("Path empty", "Please select path.")
            #select_path()
            select_when_empty()
            
        else:
            #print("message --> get_folder_check2 called 1")
            #p = file_check.read()
            #get_folder_check2(p)
            pass
            
    else:
        #print("message --> get_folder_check2 called 2")
        #get_folder_check2(p)
        pass


def select_path():
    #print('select_path called.')
    #global q
    #print('select_path called')
    global path_mainDir
    global p
    path_mainDir = filedialog.askdirectory( initialdir=p)
    
    #(path)
    path_mainDir = path_mainDir.replace("/","\\")
    #print('Dir just called')
    #print(path_mainDir)
    #if p!=path_mainDir: # Here, add a function such that, the old buttons get deleted and new buttons are added.
        #print('Difference')
        #print(f'p = {p}') # p is the previous directory path
        #print(f'path_mainDir = {path_mainDir}') # path_mainDir is the new directory path.
        #print(get_folder_check2(p)) # Prev path dict
        #print('*'*100)
        #print(get_folder_check2(path_mainDir)) # New path dict
        #Delete_Button(get_folder_check2(p))
    #print(p)
    #q = p # Storing the prev path
    p = path_mainDir # Storing the new path
    #print(p)
    if p=='':
        messagebox.showerror("Path Invalid", "Please select valid path")
        drop.configure(state="disabled")
        #print('p is empty')
        #print('select_path_called')
        #select_path()
    else:
        #print('check_folders_called')
        check_folders(p)
        
def Open_path(p):
    #global file_real
    global file_check
    file_check = open("Path.txt","w")
    file_check.write(p) # Saving the path to the file when the window is closed.  
    messagebox.showinfo("Successful", "Path successfully added")
    file_check.close()
    file_double_check = open("Path.txt","r")
    p = file_double_check.read()
    file_double_check.close()
    #print(f'p inside Open_path after reading: {p}')
    #print("Open_path --> get_folder called.") 
    get_folder(p)

def get_folder_check2(path_of_DIr):
    folder_list2 = [] # Saves the name such as BFT 1, BFT 2, BFT 3, etc.
    folder_path_list2 = [] # Saves the folder path such as 'C:\\Users\\ADMIN\\Desktop\\Test Server\\BFT_server\\BFT 1', 'C:\\Users\\ADMIN\\Desktop\\Test Server\\BFT_server\\BFT 2', etc.
    folders2 = os.listdir(path_of_DIr)
    for folder in folders2:
        folder_list2.append(folder)
        folder_path = os.path.join(path_of_DIr,folder)
        folder_path_list2.append(folder_path)
    #return Time_list_Create(folder_path_list2,folder_list2)

def get_folder(Dirpath_to_get_folder): # List of all folders in main dir, and each of their paths are stored here.
    global folder_path_list
    global folder_list
    folder_list = [] # Saves the name such as BFT 1, BFT 2, BFT 3, etc.
    folder_path_list = [] # Saves the folder path such as 'C:\\Users\\ADMIN\\Desktop\\Test Server\\BFT_server\\BFT 1', 'C:\\Users\\ADMIN\\Desktop\\Test Server\\BFT_server\\BFT 2', etc.
    folders = os.listdir(Dirpath_to_get_folder)
    for folder in folders:
        folder_list.append(folder)
        folder_path = os.path.join(Dirpath_to_get_folder,folder)
        folder_path_list.append(folder_path)
    #print("Inside get_folder")
    #print(f'folder_path_list = {folder_path_list}')
    #print(f'folder_list = {folder_list}')
    Delete_Button()
    Create_button(Time_list_Create(folder_path_list,folder_list))
    #print('get_folder called')

def Gen_Data():
    #print("HI")
    global p
    #print(f"p inside Gen_Data: {p}")
    #print("Gen_Data called.")
    #Delete_Button()
    get_folder(p)
    
def Time_list_Create(path_list,fold_list):
    global folder_list
    global folder_path_list
    dict_pos_bft={} # Dictionary with BFT name as value and its positioning for button as key, Eg., {5: BFT 1, 3: BFT 2, 1: BFT 3, 4: BFT 4, 2: BFT 5, 6: BFt 6}
    global dict_bft_logpath
    dict_bft_logpath = {} # Dictionary with BFT name as key and its log path as value.
    time_list=[]
    path_log_list = []
    #print('*'*250)
    #print("Inside Time_list_Create")
    #print(f'path_list = {path_list}')
    for path in path_list:
        #print('Inside for loop')
        #print(f'path_list = {path_list}')
        #print(f'path = {path}')
        if (os.path.isfile(os.path.join(path,'log.txt'))):
            path_log = os.path.join(path,'log.txt')
            path_log_list.append(path_log)
            time_mod = os.stat(path_log).st_mtime
            time_list.append(time_mod)
            #print(f'time_mod = {time_mod}')
        else:
            #path_list.remove(path)
            fold_list.remove(os.path.basename(path))
    #print(f'path_list 2 = {path_list}')
    #print(f'path log list = {path_log_list}')
    #print(f'fold_list = {fold_list}')
    #print(f'Time list = {time_list}')
    #Both the above lists are the same throughout the operation.

    time_sorted_list_descending = sorted(time_list,reverse=True)

    B_list = []
    for Time_ID in time_list:
        a = time_sorted_list_descending.index(Time_ID)+1 # This gives the index of the time1,time2,etc. in the list. This ID should be used to order the buttons accordingly. This ID is in INT format. eg. time 1 id is 3, then bft 1 should come at 3rd position.
        #print(a) #[5,3,1,4,2,6]
        B_list.append(a)
    #print(f'B_list = {B_list}')
    #print(fold_list)

    for i in range(len(B_list)):
        dict_pos_bft[B_list[i]] = fold_list[i]
    #print(dict_pos_bft) # {5: BFT 1, 3: BFT 2, 1: BFT 3, 4: BFT 4, 2: BFT 5, 6: BFt 6}

    for i in range(len(fold_list)):
        #print(f'fold_list 2 = {fold_list}')
        dict_bft_logpath[fold_list[i]] = path_log_list[i]
    #print(f'dict_bft_logpath = {dict_bft_logpath}')
    
    #print(B_list) # This gives the location of BFT : 1,2,3,4,5,6 in order, that should be reflected in the button.
    return dict_pos_bft
    #Create_button(dict_pos_bft)
    #destroy_button(dict_pos_bft)
    #Create_button(dict_pos_bft)
    #print('Called')
    #root.update()

def change_selected_button(button):
    #global button
    global selected_button, last_bg
    if selected_button is not None:
        selected_button.config(bg=last_bg)
    selected_button = button
    last_bg = button.cget("bg")
    button.config(bg="blue")

def ResponsiveWidget(widget, *args, **kwargs):
    bindings = {'<Enter>': {'state': 'active'},
                '<Leave>': {'state': 'normal'}}

    w = widget(*args, **kwargs)

    for (k, v) in bindings.items():
        w.bind(k, lambda e, kwarg=v: e.widget.config(**kwarg))

    return w   
    
def Create_button(dict_pos_bft):
    #global button
    global button_bft
    for key in list(dict_pos_bft.keys()):  
        button_bft = Button(Inner_frame, text=dict_pos_bft.get(key),  command=lambda key = key: BFT(dict_pos_bft.get(key)), width=13,height=3)
        button_bft.grid(row=key+1,column=0,pady=5,padx=20)
        
        buttons.append(button_bft)
    #button_bft.config(command=lambda button_bft=button_bft: change_selected_button(button_bft))
    root.after(4000,Gen_Data)
    #root.after(2000,message_)  

def Delete_Button():
    #print("Delete called")
    global buttons
    for b in buttons:
        #print(b)
        b.destroy()

def BFT(Button_name):
    #public clicked
    global dict_bft_logpath
    global AddDataPath
    AddDataPath = dict_bft_logpath.get(Button_name)
    TreeCol(AddDataPath)
    clear()
    AddData('')
    drop.configure(state="normal")
    Download_button["state"] = "normal"
    clicked.set(options[0]) # When the user clicks another button, the search dropt down will reset to the initial label (search/reset).

def clear():
    Download_button["state"] = "disable"
    # Clear the treeview
    for record in my_tree.get_children():
        my_tree.delete(record)

def Reset_search():
    clear()
    AddData('')

def Selected(event):
    
    #my_label = Label(root,text=clicked.get()).pack()
    if (clicked.get()=='<10 Days'):
        clear()
        Download_button["state"] = "normal"
        AddData('0:10')
    elif (clicked.get()=='>10 & <30 Days'):
        clear()
        Download_button["state"] = "normal"
        AddData('10:30')
    elif (clicked.get()=='>30 and <100 Days'):
        clear()
        Download_button["state"] = "normal"
        AddData('30:100')
    elif (clicked.get()=='>100 Days'):
        clear()
        Download_button["state"] = "normal"
        AddData('100:')
    else:
        Reset_search()
        Download_button["state"] = "normal"

#Generate_button  = Button(right_pane,text="Generate Data",command=Gen_Data)
#Generate_button.pack(pady=10)

# dropdown Box
options =['Search/Reset','<10 Days', '>10 & <30 Days', '>30 and <100 Days', '>100 Days']
#public clicked
clicked = StringVar()
clicked.set(options[0])

drop = OptionMenu(right_pane,clicked, *options, command=Selected)
drop.pack(pady=10)
drop.configure(state="disabled")

# ttk style Configuration
style = ttk.Style()
style.theme_use('default') # Pick a theme
style.configure("Treeview",background="#D3D3D3",foreground='Black',rowheight=25,fieldbachground='#D3D3D3') # configure Treeview colour
style.map('Treeview',background=[('selected','#347083')]) # Change selected colour
'''
def remove_many():
    x=my_tree.selection()
    for record in x:
        my_tree.delete(record)
'''

def TreeCol(path_to_log):
    global tup_headers
    # Finding Tree Columns
    df = pd.read_csv(path_to_log)
    list_headers = []
    #print(df.columns)
    for i in df.columns:
        list_headers.append(i)
        #print(i)
    #list_headers.append('Day_counter')
    tup_headers = tuple(list_headers)
    if "Date" not in tup_headers:
        messagebox.showerror("Invalid format", "Date not found")
        #Create_Tree(())
        drop.configure(state="disabled")
        Download_button["state"] = "disable"
    else:
          
        #print(f'tup_headers = {tup_headers}') 
        Create_Tree(tup_headers)
      
# Creating a treeview frame
tree_frame = Frame(right_pane)
tree_frame.pack(pady=10,padx=10,fill=BOTH, expand=1)

# Create treeview scrollbar
tree_scroll_y = Scrollbar(tree_frame)
tree_scroll_x = Scrollbar(tree_frame,orient=HORIZONTAL)
tree_scroll_y.pack(side=RIGHT,fill=Y)
tree_scroll_x.pack(side=BOTTOM,fill=BOTH)

# Creating treeview
my_tree = ttk.Treeview(tree_frame,yscrollcommand=tree_scroll_y.set,xscrollcommand=tree_scroll_x.set,selectmode="extended")
my_tree.pack(fill=BOTH, expand=1) 

# Configure scrollbar
tree_scroll_y.config(command=my_tree.yview)
tree_scroll_x.config(command=my_tree.xview)

def Create_Tree(col):
    print('Called Create_tree')
    #print(f'CReating tree, col = {col}') 
    col_list = list(col)
    col_list.append('Counter')
    col = tuple(col_list)
    #print(col)

    #Defining treeview column
    my_tree['columns'] = col

    # Format treeview column
    my_tree.column("#0",width=0,stretch=NO) # Phantom column
    for i in tup_headers:
        my_tree.column(i, anchor=CENTER,width=120,minwidth=100)
    my_tree.column("Counter",anchor=CENTER,width=0,stretch=NO)

    # Defining treeview heading
    my_tree.heading('#0',text='',anchor=CENTER)
    for i in col:
        my_tree.heading(i,text=i,anchor=CENTER)
    my_tree.heading('Counter',text='',anchor=CENTER)

def timestamp(dt):
    epoch = datetime.utcfromtimestamp(0)
    return (dt - epoch).total_seconds() * 1000.0

def AddData(val):
    
    global tup_headers
    global AddDataPath
    global Download_button
    #print(AddDataPath)
    #global tup
    df = pd.read_csv(AddDataPath) 
    for i in range(len(df)):
        count_v=0
        data_list = []
        for v in tup_headers:
            
             #print('*'*100)
            #print(v)
            #print('*'*100)
            #if v=='Start_Time':
                #print(v)
                #data=datetime.strptime(df.loc[i,v],'%d/%m/%Y %H:%M:%S') # Time in datetime object
            #elif v=='End_Time':
                #data=datetime.strptime(df.loc[i,v],'%d/%m/%Y %H:%M:%S') # Time in datetime object
            if v=='Date':
                time_date=datetime.strptime(df.loc[i,v],'%d/%m/%Y')
                currentDT = datetime.now().strftime("%Y/%m/%d")
                time_today =  datetime.strptime(currentDT,'%Y/%m/%d')
                Milli_final1 = timestamp(time_today)
                Milli_final2 = timestamp(time_date)
                Timediff_final = Milli_final1 - Milli_final2
                data_count = Timediff_final/(1000*60*60*24)
                data = df.loc[i,v]
            else:
                data = df.loc[i,v]
            
            #print(f'Data = {data}')

            if v=="Date":
                Day_count = data_count 
                if Day_count<=10:
                    Value = '0:10'
                elif (Day_count>10 and Day_count<=30):
                    Value = '10:30'
                elif (Day_count>30 and Day_count<=100):
                    Value = '30:100'
                else:
                    Value = '100:'
            data_list.append(data)
        #if count_v>=0:
        #my_tree.insert(parent='',index='end',iid=i,text='', values=data_list) # need to add a counter instead of i, which will be constat throughout fpor each row.
            #count_v=count_v+1
        #print(data_list)

        # Create Striped row tags
        my_tree.tag_configure('allrow',background="white")
        #my_tree.tag_configure('oddrow',background="white")
        #my_tree.tag_configure('evenrow',background="lightblue")
        my_tree.tag_configure('lessthan10',background='lightgreen')
        my_tree.tag_configure('greaterthan10_lessthan30',background="lightblue")
        my_tree.tag_configure('greaterthan30_lessthan100',background='orange')
        my_tree.tag_configure('greaterthan100',background='red')

        data_list.append(data_count)
        #print(data_list)
        if (val!=''):
            if(Value==val):
                data_to_append = data_list
                if (val=='0:10'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('lessthan10',))
                elif (val=='10:30'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('greaterthan10_lessthan30',))
                elif (val=='30:100'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('greaterthan30_lessthan100',))
                else:
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('greaterthan100',))

        else:
            # Adding Data to treeview column
            data_to_append = data_list # Creating a temporary tuple
            if(Value=='0:10'):
                my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('lessthan10',))
            elif(Value=='10:30'):
                my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('greaterthan10_lessthan30',))
            elif(Value=='30:100'):
                my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('greaterthan30_lessthan100',))
            else:
                my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=('greaterthan100',))

def Select_path_for_excel():
    path_to_save_file = filedialog.asksaveasfilename(initialfile = 'Untitled.xlsx',
    defaultextension=".xlsx",filetypes=[("All Files","*.*"),("Excel Sheets","*xlsx")])
    return path_to_save_file

def to_excell_():
    global tup_headers
    global path_excel_txt
    global excel_name

    cols = tup_headers # Your column headings here
    path_excel_txt = 'ExcelTemp.txt'
    excel_name = Select_path_for_excel()
    #print('*'*100)
    #print(excel_name)
    #print('*'*100)
    list_of_tree_rows = []
    with open(path_excel_txt, "w", newline='') as new_file:
        csvwriter = csv.writer(new_file, delimiter=',')
        for row_id in my_tree.get_children(): # The serial number of the rows
            #print(row_id)
            row = my_tree.item(row_id,'values')
            list_of_tree_rows.append(row)
        list_of_tree_rows = list(map(list,list_of_tree_rows))
        list_of_tree_rows.insert(0,cols)
        for row in list_of_tree_rows:
            csvwriter.writerow(row)
    save_excel()

def save_excel():
    global path_excel_txt
    global excel_name
    writer = pd.ExcelWriter(excel_name)
    df = pd.read_csv(path_excel_txt)
    df.to_excel(writer,'sheet1')
    writer.save()

Download_button = Button(right_pane,text="Download",comman=to_excell_)
Download_button.pack(pady=10)
Download_button["state"] = "disable"

browse_button = Button(Inner_frame, text='Browse', command=lambda:[Delete_Button(),select_path(),clear()])
browse_button.grid(row=0, column=2,padx=10)

#root.update()
message_()
Gen_Data()   
root.mainloop()