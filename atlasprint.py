import lasio
from tkinter import filedialog
from tkinter import *
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import openpyxl
from win32com import client
from PyPDF2 import PdfFileMerger

window = Tk()
window.configure(background='gray')
window.title("ATL PLOT EX. v1.0")
window.bind("<Escape>", lambda e: e.widget.quit())
window.geometry('1200x300')
fr = Frame(window)
bt1 = Button(fr, width=33, text='open .las', relief=RAISED, bd=6)
bt2 = Button(fr, width=33, text='save .las', relief=RAISED, bd=6)
bt3 = Button(fr, width=33, text='print md', relief=RAISED, bd=6)
bt4 = Button(fr, width=33, text='prin tvdss', relief=RAISED, bd=6)

firsttext = 'Вот мысль, которой весь я предан,\nИтог всего, что ум скопил.\nЛишь тот, кем бой за жизнь изведан,\nЖизнь и свободу заслужил.'
label1 = Label(text=firsttext, fg="#eee", bg="#333")
label1.place(relx=.2, rely=.2)

bt1.pack(side='left', )
bt2.pack(side='left', padx=2)
bt3.pack(side='left', padx=2)
bt4.pack(side='left', padx=2)

fr.pack()


def open_file(event):
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[(".las", "*.las")])

def save_las(event):
    global data
    with open(file_path, 'r') as f:
        old_data = f.read()
    # replace mnemonics txt file
    new_data = old_data.replace('MECHANICAL SPEED', 'Rate Of Penetration')
    new_data2 = new_data.replace('VERTICAL DEPTH', 'Gamma Ray')
    new_data3 = new_data2.replace('GAMMA RAY', 'True Vertical Depth Sub Sea')
    # read data usig lasio
    data = lasio.read(new_data3)

    # data.keys()
    # data to pandas data frame, rename and sort columns
    df = data.df()
    df = df.rename(columns={'GAMA': 'GR', 'GAMMA': 'GR', 'MECHSPEED': 'ROP', 'VERTICALDEPTH': 'TVDSS'})
    df = df[['GR', 'TVDSS', 'ROP']]
    global tvdss
    tvdss = df['TVDSS']
    # pandas to data, write las
    data.set_data(df)
    data.write('C:\\AtlasPrint\\TEMP\\LASFILE.las', version=2.0)
    lasname = filedialog.asksaveasfilename(title=u'save file ', filetypes=[(".las", ".las")])
    data.write(lasname, version=2.0)
    # get column DEPT(as string....) FIX IT!


def save_md(event):
    # create fig MD
    md_str = data.index

    bottom = md_str.max()
    top = md_str.min()
    bottom -= bottom % -10
    top -= top % +10
    numberdots = len(md_str) + 31
    cm = 1 / 2.54
    ysize = numberdots / 20

    fig, ax = plt.subplots(nrows=1, ncols=2, gridspec_kw={'width_ratios': [1, 2]})
    fig.set_size_inches(21 * cm, ysize * cm)
    fig.subplots_adjust(top=0.88)


    for axes in ax:
        axes.set_ylim(bottom, top)
        depth_major_ticks = np.arange(top - 2, bottom + 2, 5)
        depth_minor_ticks = np.arange(top, bottom + 1, 1)
        axes.set_yticks(depth_major_ticks)
        axes.set_yticks(depth_minor_ticks, minor=True)
        axes.set_yticks(depth_minor_ticks, minor=True)
        axes.get_xaxis().set_visible(False)
        axes.grid(which='minor', axis='y', alpha=0.5)
        axes.grid(which='major', axis='y', alpha=1)

    ax[1].set_ylabel("Measured Depth [m]", color="Black", fontsize=10, loc = 'top')

    # track 1 (MD)
    ax_GR = ax[1].twiny()
    ax_GR.set_xlim(0, 150)
    ax_GR.set_xlabel('GR [api]', color="Green", fontsize=12)
    ax_GR.plot(data['GR'], data['DEPTH' or 'DEPT'], color="Green", label='GR [api]', linewidth=1)
    ax_GR.tick_params(axis='x', colors='Green')
    ax_GR.spines['top'].set_position(('outward', 50))
    ax_GR.spines['bottom'].set_position(('outward', 50))
    major_ticks = np.arange(0, 151, 75)
    minor_ticks = np.arange(0, 151, 15)
    ax_GR.set_xticks(major_ticks)
    ax_GR.set_xticks(minor_ticks, minor=True)
    ax_GR.grid(which='minor', alpha=0.5)
    ax_GR.grid(which='major', alpha=1)
    # track 0 (MD)
    ax_ROP = ax[0].twiny()
    ax_ROP.set_xlim(100, 0)
    ax_ROP.set_xlabel('ROP [m/h]', color="Black", fontsize=12)
    ax_ROP.plot(data['ROP'], data['DEPTH'], color="Black", label='ROP [m/h]', linewidth=1)
    ax_ROP.tick_params(axis='x', colors='Black')
    ax_ROP.spines['top'].set_position(('outward', 50))
    ax_ROP.spines['bottom'].set_position(('outward', 50))
    major_ticks = np.arange(0, 101, 50)
    minor_ticks = np.arange(0, 101, 10)
    ax_ROP.set_xticks(major_ticks)
    ax_ROP.set_xticks(minor_ticks, minor=True)
    ax_ROP.grid(which='minor', alpha=0.5)
    ax_ROP.grid(which='major', alpha=1)

    fig.tight_layout()
    fig.savefig('C:\\AtlasPrint\\TEMP\\GR_PLOT_PDF_md.pdf')

    # write changes to xlsx header


    path = "C:\\AtlasPrint\\Header MD.xlsx"
    wb_obj = openpyxl.load_workbook(path.strip())
    sheet_obj = wb_obj.active
    cellThatIsToBeChanged = sheet_obj.cell(row=43, column=2)
    cellThatIsToBeChanged.value = bottom
    wb_obj.save('C:\\AtlasPrint\\Header MD.xlsx')



    #save xlsx to pdf


    excel = client.Dispatch("Excel.Application")
    sheets = excel.Workbooks.Open('C:\\AtlasPrint\\Header MD.xlsx')
    work_sheets = sheets.Worksheets[0]
    work_sheets.ExportAsFixedFormat(0, 'C:\\AtlasPrint\\TEMP\\Header MD.pdf')
    sheets.Close(True)



    # merge pdf MD


    pdfs = ['C:\\AtlasPrint\\TEMP\\Header MD.pdf', 'C:\\AtlasPrint\\TEMP\\GR_PLOT_PDF_md.pdf']

    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)
    pdfmdname = filedialog.asksaveasfilename(title=u'save file ', filetypes=[(".pdf", ".pdf")])
    merger.write(pdfmdname)
    merger.close()


def save_tvdss(event):
  # create fig TVDSS

    bottom1 = tvdss.max()
    top1 = tvdss.min()
    bottom1 -= bottom1 % -10
    top1 -= top1 % +10
    numberdots1 = len(tvdss)
    ysize1 = numberdots1 / 40
    cm = 1 / 2.54




    fig1, ax = plt.subplots(nrows=1, ncols=2, gridspec_kw={'width_ratios': [1, 2]})
    fig1.set_size_inches(21 * cm, ysize1 * cm)
    fig1.suptitle("", fontsize=20)
    fig1.subplots_adjust(top=0.88)


    for axes in ax:
        axes.set_ylim(bottom1, top1)
        depth_major_ticks = np.arange(top1 - 2, bottom1 + 2, 5)
        depth_minor_ticks = np.arange(top1 - 2, bottom1 + 2, 1)
        axes.set_yticks(depth_major_ticks)
        axes.set_yticks(depth_minor_ticks, minor=True)
        axes.set_yticks(depth_minor_ticks, minor=True)
        axes.get_xaxis().set_visible(False)
        axes.grid(which='minor', axis='y', alpha=0.5)
        axes.grid(which='major', axis='y', alpha=1)

    ax[1].set_ylabel("True Vertical Depth Sub Sea [m]", color="Black", fontsize=10, loc = 'top')
    # ax[0].set_ylabel("True Vertical Depth Sub Sea [m]", color="Black", fontsize=15, loc = 'top')
    # track 1 (TVDSS)
    ax_GR = ax[1].twiny()
    ax_GR.set_xlim(0, 150)
    ax_GR.set_xlabel('GR [api]', color="Green", fontsize=12)
    ax_GR.plot(data['GR'], data['TVDSS'], color="Green", label='GR [api]', linewidth=1)
    ax_GR.tick_params(axis='x', colors='Green')
    ax_GR.spines['top'].set_position(('outward', 50))
    ax_GR.spines['bottom'].set_position(('outward', 50))
    major_ticks = np.arange(0, 151, 75)
    minor_ticks = np.arange(0, 151, 15)
    ax_GR.set_xticks(major_ticks)
    ax_GR.set_xticks(minor_ticks, minor=True)
    ax_GR.grid(which='minor', alpha=0.5)
    ax_GR.grid(which='major', alpha=1)
    # track 0 (TVDSS)
    ax_ROP = ax[0].twiny()
    ax_ROP.set_xlim(100, 0)
    ax_ROP.set_xlabel('ROP [m/h]', color="Black", fontsize=12)
    ax_ROP.plot(data['ROP'], data['TVDSS'], color="Black", label='ROP [m/h]', linewidth=1)
    ax_ROP.tick_params(axis='x', colors='Black')
    ax_ROP.spines['top'].set_position(('outward', 50))
    ax_ROP.spines['bottom'].set_position(('outward', 50))
    major_ticks = np.arange(0, 101, 50)
    minor_ticks = np.arange(0, 101, 10)
    ax_ROP.set_xticks(major_ticks)
    ax_ROP.set_xticks(minor_ticks, minor=True)
    ax_ROP.grid(which='minor', alpha=0.5)
    ax_ROP.grid(which='major', alpha=1)
    fig1.tight_layout()

    fig1.savefig('C:\\AtlasPrint\\TEMP\\GR_PLOT_PDF_tvdss.pdf')

    #write chsnges to xlsx
    path = "C:\\AtlasPrint\\Header TVDSS.xlsx"
    wb_obj = openpyxl.load_workbook(path.strip())
    sheet_obj = wb_obj.active
    cellThatIsToBeChanged = sheet_obj.cell(row=43, column=2)
    cellThatIsToBeChanged.value = bottom1
    wb_obj.save('C:\\AtlasPrint\\Header TVDSS.xlsx')

    # conver xslasx to pdf


    excel = client.Dispatch("Excel.Application")
    sheets1 = excel.Workbooks.Open('C:\\AtlasPrint\\Header TVDSS.xlsx')
    work_sheets1 = sheets1.Worksheets[0]
    work_sheets1.ExportAsFixedFormat(0, 'C:\\AtlasPrint\\TEMP\\Header TVDSS.pdf')
    sheets1.Close(True)

    # merge pdf TVDSS
    from PyPDF2 import PdfFileMerger

    pdfs = ['C:\\AtlasPrint\\TEMP\\Header TVDSS.pdf', 'C:\\AtlasPrint\\TEMP\\GR_PLOT_PDF_tvdss.pdf']

    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)


    pdftvdssname = filedialog.asksaveasfilename(title=u'save file ', filetypes=[(".pdf", ".pdf")])
    merger.write(pdftvdssname)
    merger.close()



bt1.bind('<Button-1>', open_file)
bt2.bind('<Button-1>', save_las)
bt3.bind('<Button-1>', save_md)
bt4.bind('<Button-1>', save_tvdss)

window.mainloop()



