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

class Mahasiswa:
    def __init__(self):
        self.NIM = None
        self.Nama = None
        self.Jurusan = None



class Node:
    def __init__(self, data):
        self.data = data
        self.next = None

class LinkedList:
    def __init__(self):
        self.head = None

    def append(self, data):
        if self.head:
            cur = self.head
            while cur.next is not None:
                cur = cur.next

            cur.next = Node(data)
        else:
            self.head = Node(data)

    def __iter__(self):
        self._cur = self.head
        return self

    def __next__(self):
        if self._cur is not None:
            ret_val = self._cur.data
            self._cur = self._cur.next
            return ret_val
        else:
            raise StopIteration

daftarHadir = LinkedList()

def InsertData():
    mahasiswa = Mahasiswa()
    mahasiswa.NIM = NIMEntry.get()
    mahasiswa.Nama = namaEntry.get()
    mahasiswa.Jurusan = jurusanEntry.get()
    daftarHadir.append(mahasiswa)
    namaEntry.delete(0, END)
    NIMEntry.delete(0, END)
    jurusanEntry.delete(0, END)

def SaveData():
    global daftarHadir, informasi
    no = 1
    baris = 4
    for mahasiswa in daftarHadir:
        sheet['A' + str(baris)] = no
        DataNo = sheet['A' + str(baris)]
        DataNo.border = border
        DataNo.alignment = alignment

        sheet['B' + str(baris)] = mahasiswa.Nama
        DataNama = sheet['B' + str(baris)]
        DataNama.border = border
        DataNama.alignment = alignment

        sheet['C' + str(baris)] = mahasiswa.NIM
        DataNIM = sheet['C' + str(baris)]
        DataNIM.border = border
        DataNIM.alignment = alignment

        sheet['D' + str(baris)] = mahasiswa.Jurusan
        DataJurusan = sheet['D' + str(baris)]
        DataJurusan.border = border
        DataJurusan.alignment = alignment

        sheet['B1'] = matkulEntry.get()
        sheet['B2'] = tanggalEntry.get()

        no += 1
        baris += 1


    daftarHadir = LinkedList()
    workbook.save(filename=str(matkulEntry.get())+"_"+str(tanggalEntry.get())+".xlsx")
    informasi['text'] = "Data absen telah di save!\nNama file: "+str(matkulEntry.get())+"_"+str(tanggalEntry.get())+".xlsx"

def CreateNewData():
    global informasi
    informasi['text'] = 'Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.'
    namaEntry.delete(0, END)
    NIMEntry.delete(0, END)
    jurusanEntry.delete(0, END)
    matkulEntry.delete(0, END)
    tanggalEntry.delete(0, END)