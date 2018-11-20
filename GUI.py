# coding=utf-8
from tkinter import *
from tkinter import ttk
from generate_database import *
from generate_file import *
from tkinter import messagebox
import tkinter.filedialog as filedialog
import os

#functions for drop-down menu 

def select_door():
    door_path.set('')
    path = filedialog.askopenfilename()
    if path:
        door_path.set(path)
def select_floor():
    floor_path.set('')
    path = filedialog.askopenfilename()
    if path:
        floor_path.set(path)
def select_partition():
    partition_path.set('')
    path = filedialog.askopenfilename()
    if path:
        partition_path.set(path)
def select_insulation():
    insulation_path.set('')
    path = filedialog.askopenfilename()
    if path:
        insulation_path.set(path)
def door_find():
    door_path_find.set('')
    path = filedialog.askopenfilename()
    if path:
        door_path_find.set(path)
def floor_find():
    floor_path_find.set("")
    path = filedialog.askopenfilename()
    if path:
        floor_path_find.set(path)
def partition_find():
    partition_path_find.set("")
    path = filedialog.askopenfilename()
    if path:
        partition_path_find.set(path)
def insulation_find():
    insulation_path_find.set("")
    path = filedialog.askopenfilename()
    if path:
        insulation_path_find.set(path)
def closeWindow():
    os._exit(0)
def center_window(root, w, h):
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))
def processButton1():
    if door_path.get() != "select the path for door" and door_path.get() != "":
        try:
            generate_database(door_path.get(), "DOOR")
        except Exception as e:
            messagebox.showinfo("messageboxDOOR", str(e.args[0]))
        else:
            messagebox.showinfo("messageboxDOOR", "database generated")
    if floor_path.get() != "select the path for floor" and floor_path_find.get() != "":
        try:
            generate_database(floor_path.get(),"FLOOR")
        except Exception as e:
            messagebox.showinfo("messageboxFLOOR", str(e.args[0]))
        else:
            messagebox.showinfo("messageboxFLOOR", "database generated")
    if insulation_path.get() != "select the path for insulation" and insulation_path.get() != "":
        try:
            generate_database(insulation_path.get(), "INSULATION")
        except Exception as e:
            messagebox.showinfo("messageboxINSULATION", str(e.args[0]))
        else:
            messagebox.showinfo("messageboxINSULATION", "database generated")
    if partition_path.get() != "select the path for partition" and partition_path.get() != "":
        try:
            generate_database(partition_path.get(), "PARTITION")
        except Exception as e:
            messagebox.showinfo("messageboxPARTITION", str(e.args[0]))
        else:
            messagebox.showinfo("messageboxPARTITION", "database generated")
def processButton2():
    window.destroy()
def delete_door():
    try:
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("drop table DOOR")
        cur.close()
        conn.commit()
        conn.close()
    except:
        messagebox.showinfo("messagebox", "table DOOR does not exist")
    else:
        messagebox.showinfo("messagebox", "table DOOR deleted")
def delete_floor():
    try:
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("drop table FLOOR")
        cur.close()
        conn.commit()
        conn.close()
    except:
        messagebox.showinfo("messagebox", "table FLOOR does not exist")
    else:
        messagebox.showinfo("messagebox", "table FLOOR deleted")
def delete_insulation():
    try:
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("drop table INSULATION")
        cur.close()
        conn.commit()
        conn.close()
    except:
        messagebox.showinfo("messagebox", "table INSULATION does not exist")
    else:
        messagebox.showinfo("messagebox", "table INSULATION deleted")
def delete_partition():
    try:
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("drop table PARTITION")
        cur.close()
        conn.commit()
        conn.close()
    except:
        messagebox.showinfo("messagebox", "table PARTITION does not exist")
    else:
        messagebox.showinfo("messagebox", "table PARTITION deleted")
def print_door():
    try:
        print_table("DOOR")
    except Exception as e:
        messagebox.showinfo("messageboxDOOR", str(e.args[0]))
    else:
        messagebox.showinfo("messageboxDOOR", "table DOOR printed")
def print_floor():
    try:
        print_table("FLOOR")
    except Exception as e:
        messagebox.showinfo("messageboxFLOOR", str(e.args[0]))
    else:
        messagebox.showinfo("messageboxFLOOR", "table FLOOR printed")
def print_insulation():
    try:
        print_table("INSULATION")
    except Exception as e:
        messagebox.showinfo("messageboxINSULATION", str(e.args[0]))
    else:
        messagebox.showinfo("messageboxINSULATION", "table INSULATION printed")
def print_partition():
    try:
        print_table("PARTITION")
    except Exception as e:
        messagebox.showinfo("messageboxPARTITION", str(e.args[0]))
    else:
        messagebox.showinfo("messageboxPARTITION", "table PARTITION printed")

window = Tk()
window.title("quantitative architecture")
center_window(window,700,800)
door_path_find = StringVar()
roof_path_find = StringVar()
floor_path_find = StringVar()
insulation_path_find = StringVar()
partition_path_find = StringVar()
door_path = StringVar()
roof_path = StringVar()
floor_path = StringVar()
insulation_path = StringVar()
partition_path = StringVar()
door_path_find.set("select the path of the file to be read")
roof_path_find.set("select the path of the file to be read")
floor_path_find.set("select the path of the file to be read")
insulation_path_find.set("select the path of the file to be read")
partition_path_find.set("select the path of the file to be read")
door_path.set("select the path for door")
roof_path.set("select the path for roof")
floor_path.set("select the path for floor")
insulation_path.set("select the path for insulation")
partition_path.set("select the path for partition")
frm_data = Frame(window,height = 350,width = 700)
frm_data.place(rely = 0.05, relx = 0.1)
frm_cal = Frame(window,height = 350,width = 700)
frm_cal.place(rely = 0.52, relx = 0.1)
frm1 = Frame(frm_data,height = 20,width = 700)
frm1.place(rely = 0, relx = 0.5, x = -350)
frm2 = Frame(frm_data,height = 120,width = 700)
frm2.place(rely = 0.1, relx = 0.5, x = -350)
frm3 = Frame(frm_data,height = 20,width = 700)
frm3.place(rely = 0.8, relx = 0.5, x = -250)
frm4 = Frame(frm_cal,height = 20,width = 700)
frm4.place(rely = 0, relx = 0.5, x = -350)
frm5 = Frame(frm_cal,height = 20,width = 700)
frm5.place(rely = 0.1, relx = 0.5, x = -350)
frm6 = Frame(frm_cal,height = 20,width = 700)
frm6.place(rely = 0.8, relx = 0.5, x = -250)

label_database = Label(frm1, text = "GENERATE DATABASE VIA THE DATAS OF SELLERS")
label_database.config(font = 'Helvetica 16 bold')
label_database.pack()
label_door = Label(frm2, text = "DOOR:")
label_door.grid(row = 0, column = 0, sticky = "w")
ent_door = Entry(frm2,width=30,textvariable = door_path)
ent_door.grid(row = 1, column =0)
bt_select_door = Button(frm2, text="select path", command=select_door, font = 'Helvetica 12 bold')
bt_select_door.grid(row = 1, column =1)
bt_delete_door = Button(frm2, text = "delete door", command = delete_door, font = "Helvetica 12 bold")
bt_delete_door.grid(row = 1, column = 2)
bt_print_door = Button(frm2, text="print", command=print_door, font = 'Helvetica 12 bold')
bt_print_door.grid(row = 1, column = 3)
label_floor = Label(frm2, text = "FLOOR:")
label_floor.grid(row = 2, column = 0, sticky = "w")
ent_floor = Entry(frm2, width=30, textvariable = floor_path)
ent_floor.grid(row = 3, column = 0)
bt_select_floor = Button(frm2, text="select path", command=select_floor, font = 'Helvetica 12 bold')
bt_select_floor.grid(row = 3, column =1)
bt_delete_floor = Button(frm2, text="delete floor", command=delete_floor, font = 'Helvetica 12 bold')
bt_delete_floor.grid(row = 3, column = 2)
bt_print_floor = Button(frm2, text="print", command=print_floor, font = 'Helvetica 12 bold')
bt_print_floor.grid(row = 3, column = 3)
label_insulation = Label(frm2, text = "INSULATION:")
label_insulation.grid(row = 4, column = 0, sticky = "w")
ent_insulation = Entry(frm2,width=30,textvariable = insulation_path)
ent_insulation.grid(row = 5, column =0)
bt_select_insulation = Button(frm2, text="select path", command=select_insulation, font = 'Helvetica 12 bold')
bt_select_insulation.grid(row = 5, column =1)
bt_delete_insulation = Button(frm2, text="delete insulation", command=delete_insulation, font = 'Helvetica 12 bold')
bt_delete_insulation.grid(row = 5, column = 2)
bt_print_insulation = Button(frm2, text="print", command=print_insulation, font = 'Helvetica 12 bold')
bt_print_insulation.grid(row = 5, column = 3)
label_partition = Label(frm2, text = "PARTITION:")
label_partition.grid(row = 6, column = 0, sticky = "w")
ent_partition = Entry(frm2,width=30,textvariable = partition_path)
ent_partition.grid(row = 7, column =0)
bt_select_partition = Button(frm2, text="select path", command=select_partition, font = 'Helvetica 12 bold')
bt_select_partition.grid(row = 7, column =1)
bt_delete_partition = Button(frm2, text="delete partition", command=delete_partition, font = 'Helvetica 12 bold')
bt_delete_partition.grid(row = 7, column = 2)
bt_print_partition = Button(frm2, text="print", command=print_partition, font = 'Helvetica 12 bold')
bt_print_partition.grid(row = 7, column = 3)
bt_generate = Button(frm3, text="generate database", command=processButton1, font = 'Helvetica 14 bold')
bt_generate.grid(row=0)
label_calculate = Label(frm4, text = "WEIGHT REPORT AND PRICE SCHEDULE")
label_calculate.config(font = 'Helvetica 16 bold')
label_calculate.grid(row =0)
label_door_find = Label(frm5, text = "DOOR:")
label_door_find.grid(row = 0, column = 0, sticky = "w")
ent_door_find = Entry(frm5,width=30,textvariable = door_path_find)
ent_door_find.grid(row = 1, column =0)
bt_select_door_find = Button(frm5, text="select path", command=door_find, font = 'Helvetica 12 bold')
bt_select_door_find.grid(row = 1, column =1)
label_floor_find = Label(frm5, text = "FLOOR:")
label_floor_find.grid(row = 2, column = 0, sticky = "w")
ent_floor_find = Entry(frm5,width=30,textvariable = floor_path_find)
ent_floor_find.grid(row = 3, column =0)
bt_select_floor_find = Button(frm5, text="select path", command=floor_find, font = 'Helvetica 12 bold')
bt_select_floor_find.grid(row = 3, column =1)
label_insulation_find = Label(frm5, text = "INSULATION:")
label_insulation_find.grid(row = 4, column = 0, sticky = "w")
ent_insulation_find = Entry(frm5,width=30,textvariable = insulation_path_find)
ent_insulation_find.grid(row = 5, column =0)
bt_select_insulation_find = Button(frm5, text="select path", command=insulation_find, font = 'Helvetica 12 bold')
bt_select_insulation_find.grid(row = 5, column =1)
label_partition_find = Label(frm5, text = "PARTITION:")
label_partition_find.grid(row = 6, column = 0, sticky = "w")
ent_partition_find = Entry(frm5,width=30,textvariable = partition_path_find)
ent_partition_find.grid(row = 7, column =0)
bt_select_partition_find = Button(frm5, text="select path", command=partition_find, font = 'Helvetica 12 bold')
bt_select_partition_find.grid(row = 7, column =1)
bt_find_weight = Button(frm6, text="generate report", command =processButton2, font = 'Helvetica 14 bold')
bt_find_weight.grid(row=0)
window.protocol('WM_DELETE_WINDOW', closeWindow)
window.mainloop()

#DOOR
if door_path_find.get() != "select the path of the file to be read" and door_path_find.get() != "":
    create_table_quantity_door()
    create_table_weight_door()
    list_incomplet = []
    list_not_exist = []
    list_dict_door = excel_to_dict(door_path_find.get())
    check_not_exist = 0
    check_key_values_total = 0
    for i in range(len(list_dict_door)):
        dict_door = list_dict_door[i]
        check = check_key_values_door(dict_door)
        if check == 1:
            list_incomplet.append(dict_door)
            check_key_values_total += 1
        else:
            w_data = find_weight_door(dict_door, "database.db")
            marker = w_data[0]
            if marker == 0:
                list_not_exist.append(dict_door)
                check_not_exist += 1
            elif marker == "duo":
                rwm = "1"
                spm = "1"
                ress = w_data[1]
                list1 = []
                for i in range(len(ress)):
                    list1.append(str(ress[i][1]))
                list11 = list(set(list1))
                tup1 = tuple(list11)
                # the window for choosing RW
                window1 = Tk()
                window1.title("choose RW for door")
                center_window(window1, 250,80) # size of the window
                l1 = Label(window1, text="MARK:"+str(dict_door["MARK"])) # tag of the door
                l1.grid(column=0, row=0)
                rw = StringVar()
                rwchosen = ttk.Combobox(window1, textvariable=rw)
                rwchosen['values'] = tup1
                rwchosen.grid(column=0, row=1)
                rwchosen.current(0)
                def bt1():
                    global rwm
                    rwm = rwchosen.get()
                    window1.destroy()
                button1 = ttk.Button(window1, text="choose", command=bt1)
                button1.grid(column=0, row=2)
                window1.protocol('WM_DELETE_WINDOW', closeWindow)
                window1.mainloop()
                # the window for choosing RW
                d1 = input_rw(rwm)
                dict_merge1 = merge_dict(dict_door, d1)
                supplier_list = find_supplier_door(dict_merge1, "database.db")
                list2 = []
                for i in range(len(supplier_list)):
                    list2.append(str(supplier_list[i][0]))
                list22 = list(set(list2))
                tup2 = tuple(list22)
                window2 = Tk()
                window2.title("choose Supplier for door")
                center_window(window2, 250, 80) # size of the window
                l2 = Label(window2, text="MARK:" + str(dict_door["MARK"])) # tag of the window
                l2.grid(column=0, row=0)
                supplier = StringVar()
                supplierchosen = ttk.Combobox(window2, textvariable=supplier)
                supplierchosen['values'] = tup2
                supplierchosen.grid(column=0, row=1)
                supplierchosen.current(0)
                def bt2():
                    global spm
                    global rwm
                    spm = supplierchosen.get()
                    d2 = input_rw_sp(rwm, spm)
                    dict_merge2 = merge_dict(dict_door, d2)
                    insert_weight_level_door(dict_merge2)
                    insert_3_values_door(dict_merge2)
                    window2.destroy()
                button2 = ttk.Button(window2, text="choose", command=bt2)
                button2.grid(column=0, row=2)
                window2.protocol('WM_DELETE_WINDOW', closeWindow)
                window2.mainloop()
            else:
                insert_weight_level_door(dict_door)
                insert_3_values_door(dict_door)
    if check_key_values_total != 0:
        dict_to_csv(list_incomplet, "test_incomplet_DOOR.csv")
        top = Tk()
        top.withdraw()
        r1 = messagebox.showinfo("result_door", "the datas are incomplet and they have been written in a file")
        top.destroy()
    if check_not_exist != 0:
        dict_to_csv(list_not_exist, "test_not_exist_DOOR.csv")
        top = Tk()
        top.withdraw()
        r2 = messagebox.showinfo("result_door", "some datas not found in database and they have been written in a file")
        top.destroy()
    if check_key_values_total == 0 and check_not_exist == 0:
        ress_weight = group_weight_door()
        write_weight_door(ress_weight)
        ress = group_quantity_door()
        write_quantity_door(ress)
        top = Tk()
        top.withdraw()
        r3 = messagebox.showinfo("result_door", "succeed, weight report and price schedule generated")
        top.destroy()

if floor_path_find.get() != "select the path of the file to be read" and floor_path_find.get() != "":
    create_table_weight_floor()
    create_table_quantity_floor()
    list_incomplet = []
    list_not_exist = []
    list_dict_floor = excel_to_dict(floor_path_find.get())
    check_not_exist = 0
    check_key_values_total = 0
    for i in range(len(list_dict_floor)):
        dict_floor = list_dict_floor[i]
        check = check_key_values_floor(dict_floor)
        if check == 1:
            list_incomplet.append(dict_floor)
            check_key_values_total += 1
        else:
            w_data = find_weight_floor(dict_floor, "database.db")
            marker = w_data[0]
            if marker == 0:
                list_not_exist.append(dict_floor)
                check_not_exist += 1
            elif marker == "duo":
                rwm = "1"
                spm = "1"
                ress = w_data[1]
                list1 = []
                for i in range(len(ress)):
                    list1.append(str(ress[i][1]))
                list11 = list(set(list1))
                tup1 = tuple(list11)
                # the window for choosing RW
                window1 = Tk()
                window1.title("choose RW for floor")
                center_window(window1, 210, 100) #size of the window
                l1 = Label(window1, text=" TYPE:" + str(dict_floor["TYPE"]))  #tag1 of the floor
                l1.grid(column=0, row=0)
                l11 = Label(window1, text=" FIRERATING:" + str(dict_floor["FIRERATING"])) #tag2 of the floor
                l11.grid(column = 0, row = 1)
                rw = StringVar()
                rwchosen = ttk.Combobox(window1, textvariable=rw)
                rwchosen['values'] = tup1
                rwchosen.grid(column=0, row=2)
                rwchosen.current(0)
                def bt1():
                    global rwm
                    rwm = rwchosen.get()
                    window1.destroy()
                button1 = ttk.Button(window1, text="choose", command=bt1)
                button1.grid(column=0, row=3)
                window1.protocol('WM_DELETE_WINDOW', closeWindow)
                window1.mainloop()
                d1 = input_rw(rwm)
                dict_merge1 = merge_dict(dict_floor, d1)
                supplier_list = find_supplier_floor(dict_merge1, "database.db")
                list2 = []
                for i in range(len(supplier_list)):
                    list2.append(str(supplier_list[i][0]))
                list22 = list(set(list2))
                tup2 = tuple(list22)
                # window for choosing supplier
                window2 = Tk()
                window2.title("choose Supplier for floor")
                center_window(window2, 210, 100) #size of the window
                l2 = Label(window2, text="TYPE:" + str(dict_floor["TYPE"])) #tag1 of the floor
                l2.grid(column=0, row=0)
                l22 = Label(window2, text="FIRERATING:" + str(dict_floor["FIRERATING"])) #tag2 of the floor
                l22.grid(column = 0, row = 1)
                supplier = StringVar()
                supplierchosen = ttk.Combobox(window2, textvariable=supplier)
                supplierchosen['values'] = tup2
                supplierchosen.grid(column=0, row=2)
                supplierchosen.current(0)
                def bt2():
                    global weightduo
                    global weightduo0
                    global weightduo1
                    global weightduo2
                    global spm
                    global rwm
                    global count
                    global count0
                    global count1
                    global count2
                    spm = supplierchosen.get()
                    d2 = input_rw_sp(rwm, spm)
                    dict_merge2 = merge_dict(dict_floor, d2)
                    insert_weight_level_floor(dict_merge2)
                    insert_type_area_floor(dict_merge2)
                    window2.destroy()
                button2 = ttk.Button(window2, text="choose", command=bt2)
                button2.grid(column=0, row=3)
                window2.protocol('WM_DELETE_WINDOW', closeWindow)
                window2.mainloop()
            else:
                insert_weight_level_floor(dict_floor)
                insert_type_area_floor(dict_floor)
    if check_key_values_total != 0:
        dict_to_csv(list_incomplet, "test_incomplet_floor.csv")
        top = Tk()
        top.withdraw()
        r1 = messagebox.showinfo("result_floor", "the datas are incomplet and they have been written in a file")
        top.destroy()
    if check_not_exist != 0:
        dict_to_csv(list_not_exist, "test_not_exist_floor.csv")
        top = Tk()
        top.withdraw()
        r2 = messagebox.showinfo("result_floor", "some datas not found in database and they have been written in a file")
        top.destroy()
    if check_key_values_total == 0 and check_not_exist == 0:
        ress_weight = group_weight_floor()
        write_weight_floor(ress_weight)
        ress = group_quantity_floor()
        write_quantity_floor(ress)
        top = Tk()
        top.withdraw()
        r3 = messagebox.showinfo("result_floor", "succeed, weight report and price schedule generated")
        top.destroy()


#insulation
if insulation_path_find.get() != "select the path of the file to be read" and insulation_path_find.get() != "":
    create_table_weight_insulation()
    create_table_quantity_insulation()
    list_incomplet = []
    list_not_exist = []
    weight = 0.0
    weight0 = 0.0
    weight1 = 0.0
    weight2 = 0.0
    list_dict_insulation = excel_to_dict(insulation_path_find.get())
    check_not_exist = 0
    check_key_values_total = 0
    weightduo = 0.0
    weightduo0 = 0.0
    weightduo1 = 0.0
    weightduo2 = 0.0
    count = 0.0
    count0 = 0.0
    count1 = 0.0
    count2 = 0.0
    for i in range(len(list_dict_insulation)):
        dict_insulation = list_dict_insulation[i]
        check = check_key_values_insulation(dict_insulation)
        if check == 1:
            list_incomplet.append(dict_insulation)
            check_key_values_total += 1
        else:
            w_data = find_weight_insulation(dict_insulation, "database.db")
            marker = w_data[0]
            if marker == 0:
                list_not_exist.append(dict_insulation)
                check_not_exist += 1
            elif marker == "duo":
                rwm = "1"
                spm = "1"
                ress = w_data[1]
                list1 = []
                for i in range(len(ress)):
                    list1.append(str(ress[i][1]))
                list11 = list(set(list1))
                tup1 = tuple(list11)
                # window for choosing RW
                window1 = Tk()
                window1.title("choose RW for insulation")
                center_window(window1, 260, 110)
                l1 = Label(window1, text=" TYPE:" + str(dict_insulation["TYPE"])) # tag1 of the insulation
                l1.grid(column=0, row=0)
                l2 = Label(window1, text=" FIRERATING:" + str(dict_insulation["FIRERATING"])) #tag2 of the insulation
                l2.grid(column=0, row=1)
                rw = StringVar()
                rwchosen = ttk.Combobox(window1, textvariable=rw)
                rwchosen['values'] = tup1
                rwchosen.grid(column=0, row=2)
                rwchosen.current(0)
                def bt1():
                    global rwm
                    rwm = rwchosen.get()
                    window1.destroy()
                button1 = ttk.Button(window1, text="choose", command=bt1)
                button1.grid(column=0, row=3)
                window1.protocol('WM_DELETE_WINDOW', closeWindow)
                window1.mainloop()
                d1 = input_rw(rwm)
                dict_merge1 = merge_dict(dict_insulation, d1)
                supplier_list = find_supplier_insulation(dict_merge1, "database.db")
                list2 = []
                for i in range(len(supplier_list)):
                    list2.append(str(supplier_list[i][0]))
                list22 = list(set(list2))
                tup2 = tuple(list22)
                #window for choosing supplier
                window2 = Tk()
                window2.title("choose Supplier for insulation")
                center_window(window2, 260, 110) #size of the window
                l3 = Label(window2, text="TYPE:" + str(dict_insulation["TYPE"])) #tag1 of insulation
                l3.grid(column=0, row=0)
                l4 = Label(window2, text="FIRERATING:" + str(dict_insulation["FIRERATING"])) #tag2 of insulation
                l4.grid(column=0,row=1)
                supplier = StringVar()
                supplierchosen = ttk.Combobox(window2, textvariable=supplier)
                supplierchosen['values'] = tup2
                supplierchosen.grid(column=0, row=2)
                supplierchosen.current(0)
                def bt2():
                    global weightduo
                    global weightduo0
                    global weightduo1
                    global weightduo2
                    global spm
                    global rwm
                    global count
                    global count0
                    global count1
                    global count2
                    spm = supplierchosen.get()
                    d2 = input_rw_sp(rwm, spm)
                    dict_merge2 = merge_dict(dict_insulation, d2)
                    insert_weight_level_insulation(dict_merge2)
                    insert_firerating_volume_insulation(dict_merge2)
                    window2.destroy()
                button2 = ttk.Button(window2, text="choose", command=bt2)
                button2.grid(column=0, row=3)
                window2.protocol('WM_DELETE_WINDOW', closeWindow)
                window2.mainloop()
            else:
                insert_weight_level_insulation(dict_insulation)
                insert_firerating_volume_insulation(dict_insulation)
    if check_key_values_total != 0:
        dict_to_csv(list_incomplet, "test_incomplet_insulation.csv")
        top = Tk()
        top.withdraw()
        r1 = messagebox.showinfo("result_insulation", "the datas are incomplet and they have been written in a file")
        top.destroy()
    if check_not_exist != 0:
        dict_to_csv(list_not_exist, "test_not_exist_insulation.csv")
        top = Tk()
        top.withdraw()
        r2 = messagebox.showinfo("result_insulation",
                                 "some datas not found in database and they have been written in a file")
        top.destroy()
    if check_key_values_total == 0 and check_not_exist == 0:
        ress_weight = group_weight_insulation()
        write_weight_insulation(ress_weight)
        ress = group_quantity_insulation()
        write_quantity_insulation(ress)
        top = Tk()
        top.withdraw()
        r3 = messagebox.showinfo("result_insulation", "succeed, weight report and price schedule generated")
        top.destroy()

if partition_path_find.get() != "select the path of the file to be read" and partition_path_find.get() != "":
    create_table_weight_partition()
    create_table_quantity_partition()
    list_incomplet = []
    list_not_exist = []
    weight = 0.0
    weight0 = 0.0
    weight1 = 0.0
    weight2 = 0.0
    list_dict_partition = excel_to_dict(partition_path_find.get())
    check_not_exist = 0
    check_key_values_total = 0
    weightduo = 0.0
    weightduo0 = 0.0
    weightduo1 = 0.0
    weightduo2 = 0.0
    count = 0.0
    count0 = 0.0
    count1 = 0.0
    count2 = 0.0
    for i in range(len(list_dict_partition)):
        dict_partition = list_dict_partition[i]
        check = check_key_values_partition(dict_partition)
        if check == 1:
            list_incomplet.append(dict_partition)
            check_key_values_total += 1
        else:
            w_data = find_weight_partition(dict_partition, "database.db")
            marker = w_data[0]
            if marker == 0:
                list_not_exist.append(dict_partition)
                check_not_exist += 1
            elif marker == "duo":
                rwm = "1"
                spm = "1"
                ress = w_data[1]
                list1 = []
                for i in range(len(ress)):
                    list1.append(str(ress[i][1]))
                list11 = list(set(list1))
                tup1 = tuple(list11)
                # window for choosing RW
                window1 = Tk()
                window1.title("choose RW for partition")
                center_window(window1, 300, 120)
                l1 = Label(window1, text=" THICKNESS:" + str(dict_partition["THICKNESS"])) #tag1 of partition
                l1.grid(column=0, row=0)
                l2 = Label(window1, text=" FIRERATING:" + str(dict_partition["FIRERATING"])) #tag2 of partition
                l2.grid(column=0, row=1)
                l11 = Label(window1, text=" TYPE:" + str(dict_partition["TYPE"])) #tag3 of partition
                l11.grid(column = 0, row = 2)
                rw = StringVar()
                rwchosen = ttk.Combobox(window1, textvariable=rw)
                rwchosen['values'] = tup1
                rwchosen.grid(column=0, row=3)
                rwchosen.current(0)
                def bt1():
                    global rwm
                    rwm = rwchosen.get()
                    window1.destroy()
                button1 = ttk.Button(window1, text="choose", command=bt1)
                button1.grid(column=0, row=4)
                window1.protocol('WM_DELETE_WINDOW', closeWindow)
                window1.mainloop()
                d1 = input_rw(rwm)
                dict_merge1 = merge_dict(dict_partition, d1)
                supplier_list = find_supplier_partition(dict_merge1, "database.db")
                list2 = []
                for i in range(len(supplier_list)):
                    list2.append(str(supplier_list[i][0]))
                list22 = list(set(list2))
                tup2 = tuple(list22)
                #window for choosing supplier
                window2 = Tk()
                window2.title("choose Supplier for partition")
                center_window(window2, 300, 120) #size of the window
                l3 = Label(window2, text="THICKNESS:" + str(dict_partition["THICKNESS"])) #tag1 of partition
                l3.grid(column=0, row=0)
                l4 = Label(window2, text="FIRERATING:" + str(dict_partition["FIRERATING"])) #tag2 of partition
                l4.grid(column=0, row=1)
                l33 = Label(window2, text="TYPE:" + str(dict_partition["TYPE"])) #tag3 of partition
                l33.grid(column = 0, row = 2)
                supplier = StringVar()
                supplierchosen = ttk.Combobox(window2, textvariable=supplier)
                supplierchosen['values'] = tup2
                supplierchosen.grid(column=0, row=3)
                supplierchosen.current(0)
                def bt2():
                    global weightduo
                    global weightduo0
                    global weightduo1
                    global weightduo2
                    global spm
                    global rwm
                    global count
                    global count0
                    global count1
                    global count2
                    spm = supplierchosen.get()
                    d2 = input_rw_sp(rwm, spm)
                    dict_merge2 = merge_dict(dict_partition, d2)
                    insert_weight_level_partition(dict_merge2)
                    insert_firerating_area_partition(dict_merge2)
                    window2.destroy()
                button2 = ttk.Button(window2, text="choose", command=bt2)
                button2.grid(column=0, row=4)
                window2.protocol('WM_DELETE_WINDOW', closeWindow)
                window2.mainloop()
            else:
                insert_weight_level_partition(dict_partition)
                insert_firerating_area_partition(dict_partition)
    if check_key_values_total != 0:
        dict_to_csv(list_incomplet, "test_incomplet_partition.csv")
        top = Tk()
        top.withdraw()
        r1 = messagebox.showinfo("result_partition", "the datas are incomplet and they have been written in a file")
        top.destroy()
    if check_not_exist != 0:
        dict_to_csv(list_not_exist, "test_not_exist_partition.csv")
        top = Tk()
        top.withdraw()
        r2 = messagebox.showinfo("result_partition",
                                 "datas not found in database and they have been written in a file")
        top.destroy()
    if check_key_values_total == 0 and check_not_exist == 0:
        ress_weight = group_weight_partition()
        write_weight_partition(ress_weight)
        ress = group_quantity_partition()
        write_quantity_partition(ress)
        top = Tk()
        top.withdraw()
        r3 = messagebox.showinfo("result_partition", "succeed, weight report and price schedule generated")
        top.destroy()
