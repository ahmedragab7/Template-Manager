

import enum
#"EZC", "AUX", "HMI","PLC","SCANNER","REPEATER"
#class devices(enum.Enum):
#    EZC = 1.1
 #   AUX = 2.4
#    HMI = 3.7
 #   PLC = 7.2
 #   SCANNER = 0.3
 #   REPEATER =4.0





from enum import Enum
import tkinter as tk
from itertools import permutations
import numpy as np
list=[]

def devices_to_values(argument):
    switcher = {
        "EZC":1.3,
        "AUX":3.7,
        "HMI":1.7,
        "PLC":0.6,
        "SCANNER":2.3,
        "REPEATER":4
    }

    # get() method of dictionary data type returns
    # value of passed argument if it is present
    # in dictionary otherwise second argument will
    # be assigned as default value of passed argument
    return switcher.get(argument, "nothing")

window = tk.Tk()
devices=["EZC", "AUX", "HMI","PLC","SCANNER","REPEATER"]
variable = tk.StringVar(window)
variable.set(devices[0]) # default value

tk.OptionMenu(window,variable, *[option for option in devices]).pack()
#tk.OptionMenu(window, current, *list(CustomEnum.__members__)).pack()  # An alternative.


def permutations_logic():
    global list
    perm = permutations(list)
    for i in (perm):
        print (i)

def main_L( list_local):
    global list
    total_current=sum(list)
    num_of_dev=len(list)
    phase_current= total_current / 3
    phase1_curr=0
    phase1_dev=[]
    phase2_curr=0
    phase2_dev=[]
    phase3_curr=0
    phase3_dev=[]
    for dev in range(num_of_dev):

        taken=False
        d1=phase_current-phase1_curr+list[dev]
        d2=phase_current-phase2_curr+list[dev]
        d3=phase_current-phase3_curr+list[dev]
        if(d1>=d2 and d1>=d3 and taken==False):
            taken=True
            phase1_curr=phase1_curr+list[dev]
            phase1_dev.append(dev)
        if(d2>=d1 and d2>=d3 and taken==False):
            taken=True
            phase2_curr=phase2_curr+list[dev]
            phase2_dev.append(dev)
        if(d3>=d2 and d3>=d1 and taken==False):
            taken=True
            phase3_curr=phase3_curr+list[dev]
            phase3_dev.append(dev)

    print("ph1:",phase1_dev)
    print("ph2:",phase2_dev)
    print("ph3:",phase3_dev)

def wrong_logic():
    global list
    phase1_curr=0
    phase1_dev=[]
    phase2_curr=0
    phase2_dev=[]
    phase3_curr=0
    phase3_dev=[]
    #print([i for i in list])
    total_current=sum(list)
    num_of_dev=len(list)
    phase_current= total_current / 3
    print("total_current_pp=",phase_current)
    print("num of el=",num_of_dev)
    min_dev=min(list)
    print("min_dev=",min_dev)

    for dev in range(num_of_dev):
        p1_diff=phase_current-phase1_curr
        p2_diff=phase_current-phase2_curr
        p3_diff=phase_current-phase3_curr
        taken=False
        if (((p1_diff>p2_diff and p1_diff>p3_diff) or phase1_curr==0) and taken==False ):
            phase1_curr+=list[dev]
            phase1_dev.append(dev)
            taken=True
        if (((p2_diff>p1_diff and p2_diff>p3_diff)  or phase2_curr==0)and taken==False):
            phase2_curr+=list[dev]
            phase2_dev.append(dev)
            taken=True
        if (((p3_diff>p2_diff and p3_diff>p1_diff)  or phase3_curr==0) and taken==False):
            phase3_curr+=list[dev]
            phase3_dev.append(dev)
            taken=True




    print("phase1:",phase1_dev)
    print("phase2:",phase2_dev)
    print("phase3:",phase3_dev)



def dist_logic():
     global list
     total_current=sum(list)
     CKT=[]
     num_of_dev=len(list)
     phase_current= total_current / 3
     phase1_curr=0
     phase2_curr=0
     phase3_curr=0
     CKT_lim=16
     CKT1=[0,0,0,0]
     CKT2=[0,0,0,0]
     CKT3=[0,0,0,0]
     CKT1_dev=[]
     CKT2_dev=[]
     CKT3_dev=[]
     for dev in range(num_of_dev):
         taken=False
         if((phase1_curr+list[dev]<=phase_current+1) and taken==False):
             phase1_curr+=list[dev]
             for c in range(4):
                 if (CKT1[c]+list[dev]<=CKT_lim and taken==False):
                     CKT1[c]=+list[dev]
                     CKT1_dev.append([c,dev])
                     taken=True
         if((phase2_curr+list[dev]<=phase_current+1) and taken==False ):
             phase2_curr+=list[dev]
             for c in range(4):
                 if (CKT2[c]+list[dev]<=CKT_lim and taken==False):
                     CKT2[c]=+list[dev]
                     CKT2_dev.append([c,dev])
                     taken=True
         if((phase3_curr+list[dev]<=phase_current+1) and taken==False):
             phase3_curr+=list[dev]
             for c in range(4):
                 if (CKT3[c]+list[dev]<=CKT_lim and taken==False):
                     CKT3[c]=+list[dev]
                     CKT3_dev.append([c,dev])
                     taken=True

         p1_diff=phase_current-phase1_curr
         p2_diff=phase_current-phase2_curr
         p3_diff=phase_current-phase3_curr

         if (((p1_diff>p2_diff and p1_diff>p3_diff) or phase1_curr==0) and taken==False ):
             print("cost1")
             for c in range(4):
                 if (CKT1[c]+list[dev]<=CKT_lim and taken==False):
                     CKT1[c]=+list[dev]
                     CKT1_dev.append([c,dev])
                     taken=True
         if (((p2_diff>p1_diff and p2_diff>p3_diff)  or phase2_curr==0)and taken==False):
             print("cost2")
             for c in range(4):
                 if (CKT2[c]+list[dev]<=CKT_lim and taken==False):
                     CKT2[c]=+list[dev]
                     CKT2_dev.append([c,dev])
                     taken=True
         if (((p3_diff>p2_diff and p3_diff>p1_diff)  or phase3_curr==0) and taken==False):
             print("cost3")
             for c in range(4):
                 if (CKT3[c]+list[dev]<=CKT_lim and taken==False):
                     CKT3[c]=+list[dev]
                     CKT3_dev.append([c,dev])
                     taken=True
     print("CKT1:",CKT1_dev,"CKT2:",CKT2_dev,"CKT3:",CKT3_dev)


def ok():
    global list
    print(devices_to_values(variable.get()))
    list.append(devices_to_values(variable.get()))
    print(list)
    main_L(list)

button = tk.Button(window, text="add", command=ok)
button.pack()

window.mainloop()
