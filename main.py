from tkinter import *
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment,Border,Side
from tkinter import font as tkfont

root = Tk()
root.title("ABSENSI PERKULIAHAN")
root.resizable(width=False, height=False)
root.state("normal")
workbook = Workbook()
sheet = workbook.active

styling = tkfont.Font(family='Helvetica',weight='bold', size=15)
styling2 = tkfont.Font(family='Helvetica', size=9)

font = Font(bold=True)
border = Border(left=Side(border_style='thin',color='00000000'),
                right=Side(border_style='thin',color='00000000'),
                top=Side(border_style='thin',color='00000000'),
                bottom=Side(border_style='thin',color='00000000'))

alignment = Alignment(horizontal='center', vertical='center')


HEIGHT = 600
WIDTH = 800
canvas = Canvas(root, height=HEIGHT, width=WIDTH, bg='#8c9fde')
canvas.pack()

sheet['A1'] = "Mata Kuliah\t:"
A1 = sheet['A1']
A1.font = font
sheet['A2'] = "Tanggal Perkuliahan\t:"
A2 = sheet['A2']
A2.font = font

sheet['A3'] = "No"
A3 = sheet['A3']
A3.font = font
A3.border = border
A3.alignment = alignment

sheet['B3'] = "Nama"
B3 = sheet['B3']
B3.font = font
B3.border = border
B3.alignment = alignment

sheet['C3'] = "NIM"
C3 = sheet['C3']
C3.font = font
C3.border = border
C3.alignment = alignment

sheet['D3'] = "Jurusan"
D3 = sheet['D3']
D3.font = font
D3.border = border
D3.alignment = alignment