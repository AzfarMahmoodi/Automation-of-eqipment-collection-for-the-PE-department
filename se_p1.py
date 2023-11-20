import numpy as np
import pandas as pd
import openpyxl as op
from datetime import datetime
from tkinter import *
import re

import reg as reg

while(1):
    df1=pd.read_excel(r"C:\Users\azfar\OneDrive\Desktop\pe department\Record.xlsx")
    df2=pd.read_excel(r"C:\Users\azfar\OneDrive\Desktop\pe department\Active.xlsx")
    el=pd.read_excel(r"C:\Users\azfar\OneDrive\Desktop\pe department\EquipmentRec.xlsx")
    eq=el['Equipment name'].tolist()
    options=[]
    for item in eq:
        q=0
        q=el.iloc[el[el['Equipment name']==item].index]['Quantity'].values[0]
        if q>0: 
            s=item+': '+str(q)
            options.append(s)

    key=df2.columns
    d1={}
    def reg_check(reg):
        regex = r"^[12][0-9][A-Z]{3}[0-9]{4}$"
        if re.match(regex, reg):
            return True
        else:
            return False

    ip=Tk()
    ip.title("Issue Equipment")
    ip.geometry("340x120")
    ip.attributes('-topmost',True)
    
    clicked = StringVar()
    C=False
    D=True
    E=False
    def click_cancel():
        global C
        C=True
        ip.destroy()
            
    def click_submit():
        global reg
        global name
        global D
        reg=field1.get()
        name=field2.get()
        if(reg_check(reg)):
            D=False
        else:
            D=True
        ip.destroy()

    def click_exit():
        global E
        E=True
        ip.destroy()
    
    lbl1=Label(ip, text="Register Number: ")
    field1=Entry(ip, width=40)
    field1.insert(0, "21ABC9999")
    lbl2=Label(ip, text="Name: ")
    field2=Entry(ip, width=40)
    blank=Label(ip, text=" ")
    btn1=Button(ip, text="Submit",padx=28, command=click_submit)
    btn2=Button(ip, text="Cancel",padx=28, command=click_cancel)
    btn3=Button(ip, text="Exit",padx=28, command=click_exit)
    
    lbl1.grid(row=0,column=0)
    field1.grid(row=0,column=1,columnspan=2)
    lbl2.grid(row=1,column=0)
    field2.grid(row=1,column=1,columnspan=2)
    blank.grid(row=2,column=0)
    btn3.grid(row=4, column=1)
    btn1.grid(row=3,column=0)
    btn2.grid(row=3,column=2)
    
    ip.mainloop()

    if(E):
        break
    if(C):
        continue
    else:
        if(D):
            reger=Tk()
            reger.geometry("800x80")
            
            G=False
            def click_ok():
                global G
                G=True
                reger.destroy()
                
            em=Label(reger, text="Register number is in invalid format!!!", fg="Red",font=("Arial",25))
            Btn=Button(reger, text="Okay",padx=30, pady=15, command=click_ok)
            em.pack()
            Btn.pack()
            reger.mainloop()
            if(G):
                continue
        else:
            d1["Register Number"]=reg
            d1["Name"]=name

    rn=d1["Register Number"]
    
    arr=df1[df1["Register Number"]==rn].index.values
    if(arr.size!=0):
        ri=arr[-1]

    B=False
    if df2.isin([rn]).any().any():
        root=Tk()
        root.attributes('-topmost',True)
        root.title("Confirm status!")
    
        def click_yes():
            global A
            A=True
            root.destroy()
            return A
    
        def click_no():
            global A
            A=False
            root.destroy()
            return A
    
        blank1=Label(root, text="\t")
        blank2=Label(root, text="\t")
        myText=Label(root, text="Equipment returned?")
        myButton1=Button(root, text="Yes",padx=30,command=click_yes)
        myButton2=Button(root, text="No",padx=30,command=click_no)
        
        blank1.grid(row=0, column=0)
        blank2.grid(row=0, column=4)
        myText.grid(row=0,column=2)
        myButton1.grid(row=1, column=1)
        myButton2.grid(row=1, column=3)
        
        root.mainloop()
        if(A):
            el.loc[(el[el["Equipment name"]==df1.loc[ri,"Equipment name"]]).index,"Quantity"]+=1
            df2.drop(df2[df2["Register Number"]==rn].index,axis=0,inplace=True)
            df1.loc[ri,"Return Time"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        else:
            F=False
            def click_ok():
                global F
                F=True
                er.destroy()
                
            er=Tk()
            er.geometry("400x80")
            myLabel=Label(er,text="Return equipment first.",fg="Red", font=("Arial",25))
            myBtn=Button(er, text="Okay",padx=30, pady=15, command=click_ok)
            myLabel.pack()
            myBtn.pack()
            er.mainloop()

            if(F):
                continue
    else:
        ddm = Tk()
        ddm.title("Issue Equipment")
        ddm.geometry("310x120")
        ddm.attributes('-topmost',True)
    
        clicked = StringVar()
    
        def click_cancel():
            global B
            B=True
            ddm.destroy()
            
        def click_submit():
            eqp=clicked.get()
            d1["Equipment name"]=eqp[0:eqp.find(':')]
            el.loc[el[el["Equipment name"]==d1["Equipment name"]].index,"Quantity"]-=1
            d1["Issue Time"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ddm.destroy()
    
        myLabel1=Label(ddm, text="Name: ", font=("Arial",12,"bold"))
        myLabel2=Label(ddm, text=d1["Name"], font=("Arial",12))
        myLabel3=Label(ddm, text="Register Number: ", font=("Arial",12,"bold"))
        myLabel4=Label(ddm, text=d1["Register Number"], font=("Arial",12))
        myLabel5=Label(ddm, text="Equipment: ", font=("Arial",12,"bold"))
        drop=OptionMenu(ddm, clicked, *options)
        myButton1=Button(ddm, text="Submit",padx=28, command=click_submit)
        myButton2=Button(ddm, text="Cancel",padx=28, command=click_cancel)
            
        myLabel1.grid(row=0 ,column=0)
        myLabel2.grid(row=0 ,column=1)
        myLabel3.grid(row=1 ,column=0)
        myLabel4.grid(row=1 ,column=1)
        myLabel5.grid(row=2 ,column=0)
        drop.grid(row=2, column=1) 
        myButton1.grid(row=3 ,column=0)
        myButton2.grid(row=3 ,column=1)
    
        ddm.mainloop()
            
        if(B):
            continue
        else:
            d=pd.DataFrame(d1,index=[0])
            df2=pd.concat([df2,d],ignore_index=True)
            df1=pd.concat([df1,d],ignore_index=True)
    
    df1.to_excel(r"C:\Users\azfar\OneDrive\Desktop\pe department\Record.xlsx",index=False)
    df2.to_excel(r"C:\Users\azfar\OneDrive\Desktop\pe department\Active.xlsx",index=False)
    el.to_excel(r"C:\Users\azfar\OneDrive\Desktop\pe department\EquipmentRec.xlsx",index=False)
