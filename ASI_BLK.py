import tkinter
import openpyxl
import PIL.Image
from PIL import ImageTk
from tkinter import *
import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd
from itertools import chain
from networkx.drawing.nx_pydot import graphviz_layout
import os
os.environ["PATH"] += os.pathsep + 'C://Program Files/Graphviz/bin/'

prev_dev=-1
distance=0

class dev:
    ID=0
    prev=0
    current=0
    distance_befor_dev=0
    name=''

def prntPath(lst, node, df, lst_vst):
    for val in df.values:
        if val[0] == node:
            lst.append(val[1])
            prntPath(lst, val[1], df, lst_vst)
    if not lst[-1] in lst_vst: print('-'.join(lst))
    for l in lst: lst_vst.add(l)
    lst.pop()
    return

def find_ends_of_paths(all_b):
    ends=[]
    global main_L
    for i in range(len(all_b)):
        m=max(all_b[i])
        print('m:',m)
        name=main_L[m].name+str(m)
        ends.append(name)
        print('nameofend:',ends)
    return ends

def draw(main_l,ends_name,vd_l):
    plt.clf()
    from_l=print_main_l(main_l)[4]
    to_l=print_main_l(main_l)[5]
    print('from:',from_l)
    print('to:',to_l)
    df = pd.DataFrame({
          'From':from_l,
          'TO'  :to_l
        })

    #g = nx.DiGraph()
    g=nx.MultiDiGraph()
    g.add_nodes_from(set(chain.from_iterable(df.values)))

    for edg in df.values:
        g.add_edge(*edg)

    pos = graphviz_layout(g, prog="dot")
    nx.draw(g, pos,with_labels=True, node_shape='s')
    nx.draw_networkx_edge_labels(
    g, pos,
    #edge_labels={('EZC', 'PLC0'): 'AB'},
    font_color='red'
)
    for i in range(0,len(ends_name)):
        plt.text(pos[ends_name[i]][0],pos[ends_name[i]][1]-3,str(vd_l[i]),color='red')

    plt.draw()
    plt.show()

    lst_vst = set()
    #prntPath(['EZC'],'EZC', df, lst_vst)

def visualization(main_list):
    S_x_P=130
    S_y_P=10
    x_P=0
    y_P=0
    line_offcet=30
    distance_offset=70
    C_L=130
    tkinter.Label(window1, text ='EZC', fg ="red" ,bg ='white',font =('BOLD',11)).place(x = x_P,y =y_P)
    x_P+=C_L
    tkinter.Label(window1, text =main_list[0].name, fg ="green" ,bg ='white',font =('BOLD',9)).place(x = x_P,y =y_P)
    canvas.create_line(10+line_offcet,y_P+10,x_P,y_P+10, fill="black", width=1)
    for i in range(1,len(main_list)):
        x_P=S_x_P+C_L*(i)
        tkinter.Label(window1, text = main_list[i].name, fg ="green" ,bg ='white',font =('BOLD',9)).place(x = x_P,y =y_P)
        tkinter.Label(window1, text =main_list[i].distance_befor_dev, fg ="blue" ,bg ='white',font =('BOLD',9)).place(x = x_P-distance_offset,y =y_P)
        canvas.create_line(x_P-(C_L*(int(main_list[i].prev)+1))+line_offcet,y_P+10,x_P,y_P+10, fill="black", width=1)
        canvas.pack()


def devices_to_values(argument):
    switcher = {
        "DPS":0.015,
        "ASI_M":0.0,
        "FMD":0.11,
        "AME":0.59,
        "splitter":0.0
    }
    return switcher.get(argument, "nothing")


def all_branches(main_list):
    branches_ID=[]
    branshes=[]
    b=find_path(main_list,main_list[-1].ID)
    branshes.append(b)
    r=find_the_rest(main_list,b)
    #print('rest:',print_main_l(r)[1])
    #print('MAIN_branch:',print_main_l(b)[1])
    while(len(r)!=0):
        b=find_path(main_list,r[len(r)-1].ID)
        branshes.append(b)
        r=find_the_rest(r,b)
        #print('rest:',print_main_l(r)[1])
        #print('branch:',print_main_l(b)[1])
    for y in range(0,len(branshes)):
        print('branch',y,'->',print_main_l(branshes[y])[1])
        branches_ID.append(print_main_l(branshes[y])[1])
    calculate_all(branches_ID)
    print('branches_ID:',branches_ID)
    return branshes,branches_ID



def calculate_all(branches):
    global main_L
    current_l=[]
    distance_l=[]
    all_P_VD=[]
    print('branches:',branches)
    for i in range(0,len(branches)):
        print(branches[i])
        for id in branches[i]:

            current_l.append(float(main_L[id].current))
            distance_l.append(float(main_L[id].distance_befor_dev))

        all_P_VD.append(calculate_VD(current_l,distance_l)[1])
    print('all_vd',all_P_VD)
    return all_P_VD



def find_path(main_list,start_p):
    L1=[]
    prev_dev=main_list[start_p].prev
    prev_dev=int(prev_dev)
    #print('the_prev_ID:',prev_dev)
    L1.append(main_list[start_p])
    while(prev_dev!=-1):
        L1.append(main_list[prev_dev])
        prev_dev=int(main_list[prev_dev].prev)
    return L1

def find_the_rest(list1,list2):
    l1=len(list1)
    l2=len(list2)
    R_L=[]
    for d1 in range(l1):
        match=0
        for d2 in range(l2):
            if list1[d1].ID==list2[d2].ID :
                match=1
                break
        if (match==0):
            R_L.append(list1[d1])

    return R_L


def calculate_VD(path_destance,path_current):
    print(path_current)
    total_current=sum(path_current)
    voltage_drop_array=[]

    voltage_drop_array.append(total_current*0.5*(path_destance[0]))
    for i in range(1,len(path_current)):
        voltage_drop_array.append((total_current-path_current[i-1])*0.5*(path_destance[i]))
        total_current-=path_current[i-1]
    print("V_D:",voltage_drop_array)
    return voltage_drop_array,sum(voltage_drop_array)

def print_main_l(list):
    global main_L
    names=[]
    u_names=[]
    IDs=[]
    Currents=[]
    distances=[]
    prev_d=[]
    for i in range(len(list)):
        names.append(list[i].name)
        u_names.append(list[i].name+str(i))
        IDs.append(list[i].ID)
        Currents.append(list[i].current)
        distances.append(list[i].distance_befor_dev)
        if(i==0):
            prev_d.append('EZC')
        if (i>=1):
            print('error:',list[i].prev)
            prev_d.append(main_L[int(list[i].prev)].name+str(int(list[i].prev)))
    return names,IDs,Currents,distances,prev_d,u_names



def put_node(child_dev,distance,current,name):
    global main_L
    l=len(main_L)
    new_dev=dev()
    new_dev.distance_befor_dev=distance
    new_dev.prev=child_dev
    new_dev.current=current
    new_dev.ID=l
    new_dev.name=name
    main_L.append(new_dev)
    print(print_main_l(main_L))

def fun(en1):
    global distance
    distance=en1.widget.get()
def fun2(en2):
    global prev_dev
    prev_dev=en2.widget.get()


def ADD_DEV():
    global prev_dev
    global distance
    put_node(prev_dev,distance,devices_to_values(variable.get()),variable.get())
    M_ID=print_main_l(main_L)[1]
    b_id=all_branches(main_L)[1]
    vd_l=calculate_all(all_branches(main_L)[1])
    print('b_id:',b_id)
    ends_l=find_ends_of_paths(b_id)
    draw(main_L,ends_l,vd_l)
    #visualization(main_L)

def ADD():
    tkinter.Label(window1, text = "distance", fg ="green" ,bg ='white',font =('BOLD',17)).place(x = 60,y =60)#'username' is placed on position 00 (row - 0 and column - 0)
    tkinter.Label(window1, text = "child device:",fg ="green" ,bg ='white',font =('BOLD',17)).place(x = 60,y =120) #'username' is placed on position 00 (row - 0 and column - 0)
    en1=tkinter.Entry(window1)
    en1.place(x = 190, y = 60) # first input-field is placed on position 01 (row - 0 and column - 1)
    en1.bind("<Return>",fun)
    en2=tkinter.Entry(window1)
    en2.place(x = 190, y = 120) # first input-field is placed on position 01 (row - 0 and column - 1)
    en2.bind("<Return>",fun2)
    button_widget2 = tkinter.Button(window1,text="ADD_DEV", command =ADD_DEV,activebackground = 'white',fg = 'white',bg ='green',highlightcolor ='white',font =('BOLD',17),activeforeground ='black')
    button_widget2 = tkinter.Button(window1,text="ADD_DEV", command =ADD_DEV,activebackground = 'white',fg = 'white',bg ='green',highlightcolor ='white',font =('BOLD',17),activeforeground ='black')
    button_widget2.place(x = 10, y = 300)

if __name__ == '__main__':

    devices=["ASI_M", "DPS", "FMD", "splitter","AME"]
    main_L=[]
    window1=tkinter.Tk()
    window1.geometry("400x600")
    window1.title("ASI BLACK CALC")
    i=PIL.Image.open("asi_cal.jpg")
    im=ImageTk.PhotoImage(i)
    lable=tkinter.Label(window1,image=im)
    lable.place(x=0,y=0)
    #canvas = Canvas(window1, width=1500, height=700)
    variable = tkinter.StringVar(window1)
    variable.set(devices[0]) # default value
    tkinter.OptionMenu(window1,variable, *[option for option in devices]).place(x=10,y=30)
    button_widget = tkinter.Button(window1,text="ADD", command = ADD,activebackground = 'white',fg = 'white',bg ='green',highlightcolor ='white',font ='BOLD',activeforeground ='black')
    button_widget.place(x = 10, y = 110)


    #canvas.pack()



    window1.mainloop()


    # l1=find_path(main_L)
    # print(l1)
    # print(find_the_rest(main_L_ID,l1))
    # d=[23,12,46,70]
    # c=[1,2,3,1]
    # calculate_VD(d,c)
    # put_node(2,10,5)
    # print(main_L[-1].prev,main_L[-1].current,main_L[-1].ID)




