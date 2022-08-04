
# Import the library tkinter
from tkinter import *
from tkinter import ttk
import webbrowser
import os
import os.path
from tkinter import messagebox
from tkinter.messagebox import showinfo
from tkinter import filedialog
import pandas as pd
import datetime
from datetime import datetime
import csv
  
# Create a GUI app
root = Tk()
root.geometry('1300x600')
root.title("  Data Log Analyser")
root.iconbitmap('icon_gui.ico')

root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Create menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Creating menu item
file_menu = Menu(my_menu,tearoff=0)
my_menu.add_cascade(label="File", menu = file_menu)
file_menu.add_command(label="Open folder", command=lambda:[Delete_Button(),select_path(),clear()])
file_menu.add_separator()
file_menu.add_command(label="Quit", command=root.quit)

help_menu = Menu(my_menu,tearoff=0)
my_menu.add_cascade(label="Help", menu = help_menu)
help_menu.add_command(label = "Get Help", command=lambda: callback("https://guihelp2.coolove.repl.co/"))

#Define a callback function
def callback(url):
   webbrowser.open_new_tab(url)

# Frame left
frame_left = Frame(root,background="grey", width=355,highlightbackground="black", highlightthickness=3)
frame_left.grid(row=0,column=0,sticky="nsew")

#frame_left.grid_rowconfigure(0, weight=1)
frame_left.grid_rowconfigure(1, weight=5)
frame_left.grid_rowconfigure(2, weight=1)

# Frame browse
frame_browse = Frame(frame_left,background="grey", padx=15, pady=15)
frame_browse.grid(row=0, column=0, sticky="nsew")

# Frame button
frame_buttons = LabelFrame(frame_left, text="Recently Updated", bg="grey", fg="black",labelanchor="n", font = ("Times", "15", "bold"),highlightbackground="black",padx=10)
frame_buttons.grid(row=1,column=0, sticky="nsew") 

frame_buttons.grid_rowconfigure(1, weight=1)
frame_buttons.grid_columnconfigure(0, weight=1)

# Canvas for Frame Butoon
canvas_buttons = Canvas(frame_buttons,background="#3C6478",width=300,borderwidth=0,highlightthickness=0)
canvas_buttons.grid(row=1,column=0, sticky="nsew")
#canvas_buttons.grid_rowconfigure(1, weight=1)
#canvas_buttons.grid_columnconfigure(0, weight=1)

# Scrollbar for Canvas
canvas_scroll = ttk.Scrollbar(canvas_buttons, orient=VERTICAL, command = canvas_buttons.yview)
canvas_scroll.pack(side=RIGHT,fill=Y) 

# Configure canvas
canvas_buttons.configure(yscrollcommand=canvas_scroll.set)
canvas_buttons.bind('<Configure>',lambda e: canvas_buttons.configure(scrollregion = canvas_buttons.bbox("all")))

# Frame inside Canvas
Frame_in_canvas = Frame(canvas_buttons,background="#3C6478", width=250,borderwidth=0,highlightthickness=0)
Frame_in_canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)

# Configuring mousewheel movement of scrollbar
def _bound_to_mousewheel(event):
    canvas_buttons.bind_all("<MouseWheel>", _on_mousewheel)

def _unbound_to_mousewheel(event):
    canvas_buttons.unbind_all("<MouseWheel>")

def _on_mousewheel(event):
    canvas_buttons.yview_scroll(int(-1*(event.delta/120)), "units")

canvas_buttons.bind('<Enter>', _bound_to_mousewheel)
canvas_buttons.bind('<Leave>', _unbound_to_mousewheel)

button_bft = Button(Frame_in_canvas)
#Frame_in_canvas.bind('<Enter>',lambda e:canvas_buttons.bind("<MouseWheel>", _on_mousewheel))
canvas_buttons.create_window((0,0), window=Frame_in_canvas, anchor="nw")

# Frame dropdown
frame_dropdown = Frame(frame_left,background="grey", padx=15, pady=15)
frame_dropdown.grid(row=2, column=0, sticky="nsew")

# Frame right
frame_right = LabelFrame(root, text="Log Data", bg="#3C6478", padx=15, pady=15, width=200,labelanchor="n",borderwidth=0,highlightbackground="black", highlightthickness=3, font = ("Times", "24", "bold"))
frame_right.grid(row=0, rowspan=100, column=1, sticky="nsew")

# main 
p=''
Btn_to_del = {}
flag = False
buttons = []
selected_button = None
last_bg = None

try: # Checking if path.txt exists
    file_real = open("Path.txt","r")
    file_real.close
except: # If path.txt doesn't exist, it is created.
    file_real = open("Path.txt","w")
    file_real.close

def check_folders(check_path):
    global file_final
    global p
    folder_check_list = [] # Folder check list
    folder_path_check_list = [] # Folder path check list
    folder_path_log_check_list = [] # List containing log file path. 
    folders_check = os.listdir(check_path) # Folders

    for folder in folders_check:
        folder_check_list.append(folder)
        folder_path = os.path.join(check_path,folder)
        folder_path_check_list.append(folder_path)

    for path in folder_path_check_list:
        if (os.path.isfile(os.path.join(path,'log.txt'))):
            path_log = os.path.join(path,'log.txt')
            folder_path_log_check_list.append(path_log)
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
    else:
        Open_path(p)

def select_when_empty():
    global p
    global file_final
    path_mainDir = filedialog.askdirectory()
    path_mainDir = path_mainDir.replace("/","\\")
    p = path_mainDir
    if p=='':
        messagebox.showerror("Path Invalid", "Please select valid path")
        drop.configure(state="disabled")
        select_path()
        file_final = open("Path.txt","w")
        file_final.write(p)
        file_final.close
    else:
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
    if p=='':      
        file_check = open("Path.txt","r")
        if os.stat("Path.txt").st_size == 0:
            file_check.close
            select_when_empty()          

def select_path():
    drop.configure(state="disabled")
    Open_Report_button["state"] = "disable"
    global path_mainDir
    global p
    path_mainDir = filedialog.askdirectory( initialdir=p)
    path_mainDir = path_mainDir.replace("/","\\")
    p = path_mainDir # Storing the new path
    if p=='':
        messagebox.showerror("Path Invalid", "Please select valid path")
        drop.configure(state="disabled")
    else:
        check_folders(p)
        
def Open_path(p):
    global file_check
    file_check = open("Path.txt","w")
    file_check.write(p) # Saving the path to the file when the window is closed.  
    messagebox.showinfo("Successful", "Path successfully added")
    file_check.close()
    file_double_check = open("Path.txt","r")
    p = file_double_check.read()
    file_double_check.close()
    get_folder(p)

def get_folder_check2(path_of_DIr):
    folder_list2 = [] # Saves the name such as BFT 1, BFT 2, BFT 3, etc.
    folder_path_list2 = [] # Saves the folder path such as 'C:\\Users\\ADMIN\\Desktop\\Test Server\\BFT_server\\BFT 1', 'C:\\Users\\ADMIN\\Desktop\\Test Server\\BFT_server\\BFT 2', etc.
    folders2 = os.listdir(path_of_DIr)
    for folder in folders2:
        folder_list2.append(folder)
        folder_path = os.path.join(path_of_DIr,folder)
        folder_path_list2.append(folder_path)

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
    Delete_Button()
    Create_button(Time_list_Create(folder_path_list,folder_list))

def Gen_Data():
    global p
    get_folder(p)
    
def Time_list_Create(path_list,fold_list):
    global folder_list
    global folder_path_list
    dict_pos_bft={} # Dictionary with BFT name as value and its positioning for button as key, Eg., {5: BFT 1, 3: BFT 2, 1: BFT 3, 4: BFT 4, 2: BFT 5, 6: BFt 6}
    global dict_bft_logpath
    dict_bft_logpath = {} # Dictionary with BFT name as key and its log path as value.
    time_list=[]
    path_log_list = []
    for path in path_list:
        if (os.path.isfile(os.path.join(path,'log.txt'))):
            path_log = os.path.join(path,'log.txt')
            path_log_list.append(path_log)
            time_mod = os.stat(path_log).st_mtime
            time_list.append(time_mod)
        else:
            fold_list.remove(os.path.basename(path))

    time_sorted_list_descending = sorted(time_list,reverse=True)

    B_list = []
    for Time_ID in time_list:
        a = time_sorted_list_descending.index(Time_ID)+1 # This gives the index of the time1,time2,etc. in the list. This ID should be used to order the buttons accordingly. This ID is in INT format. eg. time 1 id is 3, then bft 1 should come at 3rd position.
        #print(a) #[5,3,1,4,2,6]
        B_list.append(a)

    for i in range(len(B_list)):
        dict_pos_bft[B_list[i]] = fold_list[i]
    #print(dict_pos_bft) # {5: BFT 1, 3: BFT 2, 1: BFT 3, 4: BFT 4, 2: BFT 5, 6: BFt 6}

    for i in range(len(fold_list)):
        dict_bft_logpath[fold_list[i]] = path_log_list[i]
    
    #print(B_list) # This gives the location of BFT : 1,2,3,4,5,6 in order, that should be reflected in the button.
    return dict_pos_bft

def change_selected_button(button):
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
        button_bft = Button(Frame_in_canvas, text=dict_pos_bft.get(key),  command=lambda key = key: BFT(dict_pos_bft.get(key)), width=10,height=3)
        button_bft.grid(row=key+1,column=0,pady=5,padx=20)
        
        buttons.append(button_bft)
    root.after(4000,Gen_Data)

def Delete_Button():
    global buttons
    for b in buttons:
        b.destroy()

def BFT(Button_name):
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
    Open_Report_button["state"] = "disable"
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

# dropdown Box
options =['Search/Reset','<10 Days', '>10 & <30 Days', '>30 and <100 Days', '>100 Days']
#public clicked
clicked = StringVar()
clicked.set(options[0])

drop = OptionMenu(frame_dropdown,clicked, *options, command=Selected)
drop.pack(pady=10)
drop.configure(state="disabled")

# ttk style Configuration
style = ttk.Style() 
style.theme_use('default') # Pick a theme
style.configure("Treeview",background="#D3D3D3",foreground='Black',rowheight=25,fieldbachground='#D3D3D3') # configure Treeview colour
style.map('Treeview',background=[('selected','#347083')]) # Change selected colour

def TreeCol(path_to_log):
    global tup_headers
    # Finding Tree Columns
    df = pd.read_csv(path_to_log)
    list_headers = []
    for i in df.columns:
        list_headers.append(i)
    tup_headers = tuple(list_headers)
    if "Date" not in tup_headers:
        messagebox.showerror("Invalid format", "Date not found")
        drop.configure(state="disabled")
        Download_button["state"] = "disable"
        Open_Report_button["state"] = "disable"
    else:
        Create_Tree(tup_headers)

def Open_report():
    global Report_path
    try:
        os.startfile(Report_path)
    except:pass
    

# Creating a treeview frame
tree_frame = Frame(frame_right)
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
    col_list = list(col)
    col_list.append('Counter')
    col = tuple(col_list)

    #Defining treeview column
    my_tree['columns'] = col

    # Format treeview column
    my_tree.column("#0",width=0,stretch=NO) # Phantom column
    for i in tup_headers:
        if i=="Report Path":
           my_tree.column(i, anchor=CENTER,width=500,minwidth=200)
        else: 
            my_tree.column(i, anchor=CENTER,width=120,minwidth=120)
    my_tree.column("Counter",anchor=CENTER,width=0,stretch=NO)

    # Defining treeview heading
    my_tree.heading('#0',text='',anchor=CENTER)
    for i in col:
        my_tree.heading(i,text=i,anchor=CENTER)
    my_tree.heading('Counter',text='',anchor=CENTER)

def timestamp(dt):
    epoch = datetime.utcfromtimestamp(0)
    return (dt - epoch).total_seconds() * 1000.0

'''
def new(event):
    
'''
def OnSingleClick(event):
    global Report_path
    try:
        item = my_tree.identify('item',event.x,event.y)
        #print("you clicked on", my_tree.item(item,"values"))
        val_of_row = my_tree.item(item,"values")
        for i,v in enumerate(tup_headers):
            if v=="Report Path":
                Report_path = val_of_row[i]
        if (os.path.exists(Report_path)):
            Open_Report_button["state"] = "normal"
        else:
            Open_Report_button["state"] = "disable"
    except:
        pass


def OnDoubleClick(event):
    try:
        item = my_tree.identify('item',event.x,event.y)
        #print("you clicked on", my_tree.item(item,"values"))
        val_of_row = my_tree.item(item,"values")
        for i,v in enumerate(tup_headers):
            if v=="Report Path":
                Report_path = val_of_row[i]
        if (os.path.exists(Report_path)):
            Open_report()  
    except:
        pass
    #messagebox.showerror("File Invalid", "FIle not found")
            
def AddData(val):
    global tup_headers
    global AddDataPath
    global Download_button
    df = pd.read_csv(AddDataPath) 
    for i in range(len(df)):
        count_v=0
        data_list = []
        for v in tup_headers:
            if v=='Date':
                time_date=datetime.strptime(df.loc[i,v],'%d/%m/%Y')
                currentDT = datetime.now().strftime("%Y/%m/%d")
                time_today =  datetime.strptime(currentDT,'%Y/%m/%d')
                Milli_final1 = timestamp(time_today)
                Milli_final2 = timestamp(time_date)
                Timediff_final = Milli_final1 - Milli_final2
                data_count = Timediff_final/(1000*60*60*24)
                data = df.loc[i,v]
            elif v=="Report Path":
                data = df.loc[i,v]
                #print("going to print path")
                #print(data)
            else:
                data = df.loc[i,v]

            if v=="Result":
                try:
                    if data=="PASS":
                        status = "pass"
                    elif data=="FAIL":
                        status = "fail"
                    else:
                        status = "unknown"
                except:
                    pass
            
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

        # Create Striped row tags
        my_tree.tag_configure('allrow',background="white")
        #my_tree.tag_configure('oddrow',background="white")
        #my_tree.tag_configure('evenrow',background="lightblue")
        #my_tree.tag_configure('lessthan10',background='lightgreen')
        #my_tree.tag_configure('greaterthan10_lessthan30',background="lightblue")
        #my_tree.tag_configure('greaterthan30_lessthan100',background='orange')
        #my_tree.tag_configure('greaterthan100',background='red')
        my_tree.tag_configure('pass',background='lightgreen')
        my_tree.tag_configure('fail',background='red')
        my_tree.tag_configure('unknown',background='violet')

        data_list.append(data_count)

        if (status=="pass"):
            if (val!=''):
                if(Value==val):
                    data_to_append = data_list
                    if (val=='0:10'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=("pass",))
                    elif (val=='10:30'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",) )
                    elif (val=='30:100'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",) )
                    else:
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",) )
            else:
                # Adding Data to treeview column
                data_to_append = data_list # Creating a temporary tuple
                if(Value=='0:10'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",) )
                elif(Value=='10:30'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",) )
                elif(Value=='30:100'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",))
                else:
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("pass",))

        elif(status=="fail"):
            if (val!=''):
                if(Value==val):
                    data_to_append = data_list
                    if (val=='0:10'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=("fail",))
                    elif (val=='10:30'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",) )
                    elif (val=='30:100'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",) )
                    else:
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",) )
            else:
                # Adding Data to treeview column
                data_to_append = data_list # Creating a temporary tuple
                if(Value=='0:10'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",) )
                elif(Value=='10:30'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",) )
                elif(Value=='30:100'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",))
                else:
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("fail",))
        else:
            if (val!=''):
                if(Value==val):
                    data_to_append = data_list
                    if (val=='0:10'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append, tags=("unknown",))
                    elif (val=='10:30'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",) )
                    elif (val=='30:100'):
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",) )
                    else:
                        my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",) )
            else:
                # Adding Data to treeview column
                data_to_append = data_list # Creating a temporary tuple
                if(Value=='0:10'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",) )
                elif(Value=='10:30'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",) )
                elif(Value=='30:100'):
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",))
                else:
                    my_tree.insert(parent='',index='end',iid=i,text='', values=data_to_append,tags=("unknown",))


    my_tree.bind("<Double-1>", OnDoubleClick)
    my_tree.bind("<Button-1>", OnSingleClick)

def Select_path_for_excel():
    path_to_save_file = filedialog.asksaveasfilename(initialfile = 'Untitled.xlsx',
    defaultextension=".xlsx",filetypes=[("All Files","*.*"),("Excel Sheets","*xlsx")])
    return path_to_save_file

def to_excell_():
    global tup_headers
    global path_excel_txt
    global excel_name

    cols = tup_headers # Your column headings here
    #print(f'Cols = {cols}')
    path_excel_txt = 'ExcelTemp.txt'
    excel_name = Select_path_for_excel()
    list_of_tree_rows = []
    with open(path_excel_txt, "w", newline='') as new_file:
        csvwriter = csv.writer(new_file, delimiter=',')
        for row_id in my_tree.get_children(): # The serial number of the rows
            #print(f'row_id = {row_id}')
            row = my_tree.item(row_id,'values')
            last_element_index = len(row)-1
            row = row[:last_element_index]
            #print(f'row = {row}')
            list_of_tree_rows.append(row)
        list_of_tree_rows = list(map(list,list_of_tree_rows))
        #print(f'list of tree rows = {list_of_tree_rows}')
        list_of_tree_rows.insert(0,cols)
        #print(f'list of tree rows after insert = {list_of_tree_rows}')
        for row in list_of_tree_rows:
            #print(f'row in list of tree rows = {row}')
            csvwriter.writerow(row)
    save_excel()

def save_excel():
    global path_excel_txt
    global excel_name
    writer = pd.ExcelWriter(excel_name)
    df = pd.read_csv(path_excel_txt)
    df.to_excel(writer,'sheet1')
    writer.save()

Open_Report_button = Button(frame_right,text="Open Report",command=Open_report)
Open_Report_button.pack(side=LEFT,pady=10,padx=(550,0))
Open_Report_button["state"] = "disable"

Download_button = Button(frame_right,text="Download",command=to_excell_)
Download_button.pack(side=LEFT,pady=10,padx=10)
Download_button["state"] = "disable"

browse_button = Button(frame_browse, text='Browse', command=lambda:[Delete_Button(),select_path(),clear()])
browse_button.pack(pady=10)

message_()
Gen_Data()   
root.mainloop()