import pyautogui as pag
from time import sleep,strptime
from datetime import datetime,timedelta
import pandas as pd

# sleep(4)
# file = pag.locateCenterOnScreen("file.png")
# pag.moveTo(file)

# pag.click()



def ppdateplan(plan_date):
	pos = pag.locateCenterOnScreen("packingplandatepymac.png")
	pag.moveTo(pos.x + 150 ,pos.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(plan_date)

def ppdateplan2(plan_date2):
	pos1 = pag.locateCenterOnScreen("packingplandatepymac.png")
	pag.moveTo(pos1.x + 255 ,pos1.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(plan_date2)

def order_no(orderno):
	order = pag.locateCenterOnScreen("ordernopymac.png")
	pag.moveTo(order.x + 130 ,order.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(orderno)

def lot_no(lotno):
	lot = pag.locateCenterOnScreen("lotnopymac.png")
	pag.moveTo(lot.x + 82 ,lot.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(lotno)

def lot_no2(lotno2):
	lot2 = pag.locateCenterOnScreen("lotnopymac.png")
	pag.moveTo(lot2.x + 140 ,lot2.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(lotno)

def case_no(caseno):
	case = pag.locateCenterOnScreen("casenopymac.png")
	pag.moveTo(case.x + 82 ,case.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(caseno)

def case_no2(caseno2):
	case2 = pag.locateCenterOnScreen("casenopymac.png")
	pag.moveTo(case2.x + 140 ,case2.y )
	pag.click()
	pag.press("backspace", presses=10)
	pag.write(caseno2)

def tekan_button(tekanbutton):
	tekan = pag.locateCenterOnScreen("printpymac.png")
	pag.moveTo(tekan)
	pag.click()

def tekan_print(tekanprint):
	printe = pag.locateCenterOnScreen("logoprintpymac.png")
	pag.moveTo(printe)
	pag.click()

def tekan_printlagi(tekanprintlagi):
	printelagi = pag.locateCenterOnScreen("okepymac.png")
	pag.moveTo(printelagi)
	pag.click()

def tekan_back(tekanback):
	backpymac = pag.locateCenterOnScreen("backpymac.png")
	pag.moveTo(backpymac)
	pag.click()

def input_tanggal():
	hariini = (datetime.now().strftime("%Y%m%d"))
	tgl = pag.prompt(title="Masukan Tanggal", default=hariini)
	return tgl

sleep(5)
logo = pag.locateCenterOnScreen("b.png")
pag.moveTo(logo)
pag.click(clicks=2, interval=0.25)

sleep(6)
username = pag.locateCenterOnScreen("username.png")
pag.moveTo(username)
pag.moveTo(username.x + 250, username.y)
pag.click()
pag.press("backspace", presses=10)
pag.write("VU05221")

passwordd = pag.locateCenterOnScreen("passwordd.png")
pag.moveTo(passwordd)
pag.moveTo(passwordd.x + 250, passwordd.y)
pag.click()
pag.press("backspace", presses=10)
pag.write("yamaha6*")

loginawal = pag.locateCenterOnScreen("loginawal.png")
pag.moveTo(loginawal)
pag.click()

sleep(6)
ckdwj = pag.locateCenterOnScreen("ckdwj.png")
pag.moveTo(ckdwj)
pag.click(clicks=2, interval=0.25)

sleep(10)
listprint = pag.locateCenterOnScreen("list.png")
pag.moveTo(listprint)
pag.click(clicks=3, interval=0.25)

sleep(4)
if __name__ == "__main__":
	tglawal = input_tanggal()
	# print(tglawal)

	df = pd.read_excel("FAS.xlsx",engine="openpyxl", sheet_name="FAS")
	sleep(2)
	for index, row in df.iterrows():
		# besok = (datetime.now()+timedelta(1)).strftime("%Y%m%d")
		# packingplandate=row[besok]
		# packingplandate=row["packingplandate"].strftime("%Y%m%d")
		# tgll=strptime(row["packingplandate"],format="%Y%m%d")
		tgll=datetime.strptime(row["packingplandate"], "%Y-%m-%d")
		packingplandate=tgll.strftime("%Y%m%d")
		# print(packingplandate)
		if packingplandate == tglawal:
			casemarkno=row["casemarkno"]
			lotno=str(row["lotno"])
			ckdsetno=row["ckdsetno"]

			ppdateplan(plan_date=packingplandate)
			ppdateplan2(plan_date2=packingplandate)
			order_no(orderno=casemarkno)
			lot_no(lotno=lotno)
			lot_no2(lotno2=lotno)

			df_mc = pd.read_excel("mp.xlsx",engine="openpyxl",sheet_name=ckdsetno)
			for idx, baris in df_mc.iterrows():
				for i in range(1,9):
					cn = str(baris[f"case_no{i}"])
					if cn != "nan":
						case_no(caseno=cn)
						case_no2(caseno2=cn)
						tekan_button(tekanbutton=cn)
						sleep(3)
						tekan_print(tekanprint=cn)
						sleep(3)
						tekan_printlagi(tekanprintlagi=cn)
						sleep(6)
						tekan_back(tekanback=cn)
						sleep(3)
						print(cn)