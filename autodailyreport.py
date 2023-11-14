import pyautogui as pg
import  webbrowser as wb
import time

ORG_NAME = 'SWP'
DATE_1 = '01/09/2023'
DATE_2 = '30/09/2023'
url = 'http://192.168.1.160/reportoracle/frm_DialyReportByCustomer.aspx'

wb.open(url) #เปิด Web ที่ต้องการ
time.sleep(2)


# Select Organization: SWP
pg.position(3911,465) #(4007,465)
time.sleep(2)
pg.click(3911,465)
time.sleep(2)
pg.moveTo(x=3911,y=635,duration=2)
#pg.press('backspace',presses=10,interval=0.1)
#pg.write(ORG_NAME,interval=0.1)
pg.click(3911,635,interval=0.25)
time.sleep(2)

#data1
pg.moveTo(x=3938,y=423)
time.sleep(2)
pg.click(3938,423)
pg.press('backspace',presses=10,interval=0.1)
time.sleep(2)
pg.write(DATE_1,interval=0.1)
time.sleep(2)


#date2
pg.moveTo(x=4487,y=425)
time.sleep(2)
pg.click(4487,425)
pg.press('backspace',presses=10,interval=0.1)
time.sleep(2)
pg.write(DATE_2,interval=0.1)
time.sleep(2)



#click button ค้นหาข้อมูล
time.sleep(2)
pg.moveTo(x=4752,y=483,duration=2)
pg.click(4752,483)
time.sleep(2)


#click button แสดงรายละเอียดทั้งหมด
pg.moveTo(x=4035,y=1427,duration=2)
time.sleep(2)
pg.click(4035,y=1427)
time.sleep(2)


#click button Export to Excel
pg.moveTo(x=3722,y=512,duration=2)
time.sleep(5)
pg.click(x=3722,y=512)

#Close Program
pg.moveTo(x=5718,y=24,duration=2)
time.sleep(2)
pg.click(x=5718,y=24)
