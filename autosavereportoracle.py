
from tkinter import *
from tkinter import ttk,messagebox
import pyautogui as pg
import  webbrowser as wb
import time

url = 'http://192.168.1.160/reportoracle/frm_DialyReportByCustomer.aspx'
def startSaveData():
    
    # Select Organization:
    if orgname_entry.get() == '':
        messagebox.showerror('แจ้งเตือน','กรุณาป้อนข้อมูล ORG')
    elif orgname_entry.get() == 'SWP':
        pg.position(3911,465) #(4007,465)
        time.sleep(2)
        pg.click(3911,465)
        time.sleep(2)
        pg.moveTo(x=3911,y=635,duration=2)
        #pg.press('backspace',presses=10,interval=0.1)
        #pg.write(ORG_NAME,interval=0.1)
        pg.click(3911,635,interval=0.25)
        time.sleep(2)
    elif orgname_entry.get() == 'Suksawat':
        pg.position(3911,465) #(4007,465)
        time.sleep(2)
        pg.click(3911,465)
        time.sleep(2)
        pg.moveTo(x=3911,y=582,duration=2)
        #pg.press('backspace',presses=10,interval=0.1)
        #pg.write(ORG_NAME,interval=0.1)
        pg.click(3911,582,interval=0.25)
        time.sleep(2)     

    else:
        wb.open(url) #เปิด Web ที่ต้องการ
        time.sleep(2)


        #data1
        pg.moveTo(x=3938,y=423)
        time.sleep(2)
        pg.click(3938,423)
        pg.press('backspace',presses=10,interval=0.1)
        time.sleep(2)
        pg.write(v_datefrom.get(),interval=0.1)
        time.sleep(2)

        #date2
        pg.moveTo(x=4487,y=425)
        time.sleep(2)
        pg.click(4487,425)
        pg.press('backspace',presses=10,interval=0.1)
        time.sleep(2)
        pg.write(v_dateto.get(),interval=0.1)
        time.sleep(2)

        #click button ค้นหาข้อมูล
        time.sleep(2)
        pg.moveTo(x=4752,y=483,duration=2)
        pg.click(4752,483)
        time.sleep(2)

        #click button แสดงรายละเอียดทั้งหมด
        pg.moveTo(x=4035,y=1427,duration=2)
        time.sleep(2)
        pg.click(4035,1427)
        time.sleep(2)

        #click button Export to Excel
        pg.moveTo(x=3722,y=512,duration=2)
        time.sleep(10)
        pg.click(3722,512)
        time.sleep(2)

        # Close the Program
        pg.moveTo(x=5726, y=21)
        time.sleep(5)
        pg.click(5726,21)
        time.sleep(2)
        messagebox.showinfo('แจ้งข้อมูล','ดำเนินการบันทึกเรียบร้อย')
    
GUI = Tk()
GUI.title('Robotic Process Automation')
GUI.geometry('900x700')

title_label = Label(GUI,text='Products Output & Waste Sorting Customer ',font=('tahoma',16))
title_label.pack(pady=10)

rpa_label = Label(GUI,text='Robotic Process Automation ',font=('tahoma',16))
rpa_label.pack(pady=5)


v_url = StringVar()
v_url.set(url)
url_label = Entry(GUI,textvariable=v_url,font=('tahoma',12),width=60)
url_label.pack(pady=10)

frame1 = Frame(GUI)
frame1.pack()

label_frame = LabelFrame(GUI,text='รายงานผลการผลิตชิ้นงาน',font=('tahoma',12))
label_frame.place(x=100,y=200,width=600,height=400)

orgname_label = Label(label_frame, text='ORG Name:',font=('tahoma',14))
orgname_label.grid(row=0,column=0,pady=25,padx= 20,sticky=W)
dateform_label = Label(label_frame, text='Date From:',font=('tahoma',14))
dateform_label.grid(row=1,column=0,pady=25,padx= 20,sticky=W)
dateto_label = Label(label_frame, text='Date To:',font=('tahoma',14))
dateto_label.grid(row=2,column=0,pady=25,padx=20,sticky=W)

org_list = ('SWP','Suksawat')
orgname_entry = ttk.Combobox(label_frame,width=15,values=org_list,font=('tahoma',14))
orgname_entry.grid(row=0,column=1,padx=20,pady=15,sticky=W)

v_datefrom = StringVar()
datefrom_entry = Entry(label_frame,textvariable=v_datefrom,font=('tahoma',14))
datefrom_entry.grid(row=1,column=1,padx=20,pady=15,sticky=W)

v_dateto = StringVar()
datefrom_entry = Entry(label_frame,textvariable=v_dateto,font=('tahoma',14))
datefrom_entry.grid(row=2,column=1,padx=20,pady=5,sticky=W)

start_button = Button(label_frame,text='Start Process',font=('tahoma',14),command=startSaveData)
start_button.grid(row=3,column=0,padx=10,pady=30)

exit_button = Button(label_frame,text='Exit',font=('tahoma',14),command=exit)
exit_button.grid(row=3,column=1,padx=10,pady=30,ipadx=60)
GUI.mainloop()