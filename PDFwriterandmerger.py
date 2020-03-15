from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pandas
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox


class PDF_merger_and_writer:

    def __init__(self):
        gui = Tk()
        self.facturi_folder_path = StringVar()
        self.note_de_ambalaj_folder_path = StringVar()
        self.corectie_1_folder_path = StringVar()
        self.corectie_2_folder_path = StringVar()
        self.output_folder_path = StringVar()
        self.excel_file_path = StringVar()
        gui.geometry('400x200')
        gui.title('Pdf merger')
        #facturi_folder
        facturi_folder_label = Label(gui, text='folderul cu Facturi:')
        facturi_folder_label.grid(row=0, column=0)
        facturi_folder_entry = Entry(gui, textvariable=self.facturi_folder_path)
        facturi_folder_entry.grid(row=0, column=1)
        facturi_folder_browse = ttk.Button(gui, text='Browse', command=self.get_facturi_folder_path)
        facturi_folder_browse.grid(row=0, column=2)
        #note de amablaj folder
        note_de_ambalaj_folder_label = Label(gui, text='folderul cu note_de_ambalaj:')
        note_de_ambalaj_folder_label.grid(row=1, column=0)
        note_de_ambalaj_folder_entry = Entry(gui, textvariable=self.note_de_ambalaj_folder_path)
        note_de_ambalaj_folder_entry.grid(row=1, column=1)
        note_de_ambalaj_folder_browse = ttk.Button(gui, text='Browse', command=self.get_note_de_ambalaj_folder_path)
        note_de_ambalaj_folder_browse.grid(row=1, column=2)
        #corectie 1 folder
        corectie_1_folder_label = Label(gui, text='folderul cu corectie_1:')
        corectie_1_folder_label.grid(row=2, column=0)
        corectie_1_folder_entry = Entry(gui, textvariable=self.corectie_1_folder_path)
        corectie_1_folder_entry.grid(row=2, column=1)
        corectie_1_folder_browse = ttk.Button(gui, text='Browse', command=self.get_corectie_1_folder_path)
        corectie_1_folder_browse.grid(row=2, column=2)
        #corectie 2 folder
        corectie_2_folder_label = Label(gui, text='folderul cu corectie_2:')
        corectie_2_folder_label.grid(row=3, column=0)
        corectie_2_folder_entry = Entry(gui, textvariable=self.corectie_2_folder_path)
        corectie_2_folder_entry.grid(row=3, column=1)
        corectie_2_folder_browse = ttk.Button(gui, text='Browse', command=self.get_corectie_2_folder_path)
        corectie_2_folder_browse.grid(row=3, column=2)
        #output folder
        output_folder_label = Label(gui, text='folderul de output:')
        output_folder_label.grid(row=4, column=0)
        output_folder_entry = Entry(gui, textvariable=self.output_folder_path)
        output_folder_entry.grid(row=4, column=1)
        output_folder_browse = ttk.Button(gui, text='Browse', command=self.get_output_folder_path)
        output_folder_browse.grid(row=4, column=2)
        #excel_file
        output_folder_label = Label(gui, text='Excel report:')
        output_folder_label.grid(row=5, column=0)
        output_folder_entry = Entry(gui, textvariable=self.excel_file_path)
        output_folder_entry.grid(row=5, column=1)
        output_folder_browse = ttk.Button(gui, text='Browse', command=self.get_excel_file_path)
        output_folder_browse.grid(row=5, column=2)
        #Merge and write button
        merge_button =  ttk.Button(gui, text='Merge PDFs', command=self.merge_the_pdfs)
        merge_button.grid(row=7, column=1)
        write_button = ttk.Button(gui, text='Write NIR on invoices', command=self.write_on_invoices)
        write_button.grid(row=7, column=0)
        #
        gui.grid_rowconfigure(6, minsize=10)
        gui.mainloop()

    def get_facturi_folder_path(self):
        self.facturi_folder_selected = filedialog.askdirectory()
        self.facturi_folder_path.set(self.facturi_folder_selected)

    def get_note_de_ambalaj_folder_path(self):
        self.note_de_ambalaj_folder_selected = filedialog.askdirectory()
        self.note_de_ambalaj_folder_path.set(self.note_de_ambalaj_folder_selected)

    def get_corectie_1_folder_path(self):
        self.corectie_1_folder_selected = filedialog.askdirectory()
        self.corectie_1_folder_path.set(self.corectie_1_folder_selected)

    def get_corectie_2_folder_path(self):
        self.corectie_2_folder_selected = filedialog.askdirectory()
        self.corectie_2_folder_path.set(self.corectie_2_folder_selected)

    def get_output_folder_path(self):
        self.output_folder_selected = filedialog.askdirectory()
        self.output_folder_path.set(self.output_folder_selected)

    def get_excel_file_path(self):
        self.excel_file_selected = filedialog.askopenfilename()
        self.excel_file_path.set(self.excel_file_selected)
        self.df = pandas.read_excel(self.excel_file_selected, converters={'Factura': str, 'Nota ambalaj': str,
                                                                          'NIR': str, 'Corectie 1': str,
                                                                          'Corectie 2': str})

    @staticmethod
    def write_text(pedeef, numar):
        packet = io.BytesIO()
        # create a new PDF with Reportlab
        can = canvas.Canvas(packet, pagesize=letter)
        #search for way to bold the text
        can.drawString(x=250, y=560, text=numar)
        can.save()
        #move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        # read your existing PDF
        existing_pdf = PdfFileReader(open(pedeef, "rb"))
        output = PdfFileWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        # finally, write "output" to a real file
        outputStream = open(pedeef[:-4]+'_'+numar[5:]+'.pdf', "wb")
        output.write(outputStream)
        outputStream.close()

    @staticmethod
    def merge_pdfs(paths, output):
        pdf_writer = PdfFileWriter()
        for path in paths:
            pdf_reader = PdfFileReader(path)
            for page in range(pdf_reader.getNumPages()):
                # Add each page to the writer object
                pdf_writer.addPage(pdf_reader.getPage(page))
            with open(output, 'wb') as out:
                pdf_writer.write(out)

    def write_on_invoices(self):
        for i in self.df.iterrows():
            try:
                self.write_text(pedeef=str(self.facturi_folder_selected) + '/' + str(i[1]['Factura']) + '.pdf',
                                numar='NIR: '+str(i[1]['NIR']))
            except FileNotFoundError:
                pass
        messagebox.showinfo(title='NIR writer', message='NIR written on invoices')

    def merge_the_pdfs(self):
        for i in self.df.iterrows():
            factura_si_nir = str(i[1]['Factura']) + '_' + str(i[1]['NIR'])
            nota_de_ambalaj = str(i[1]['Nota ambalaj'])
            cor1 = str(i[1]['Corectie 1'])
            cor2 = str(i[1]['Corectie 2'])
            files = [str(self.facturi_folder_selected) + '/' + factura_si_nir + '.pdf',
                     str(self.note_de_ambalaj_folder_selected) + '/' + nota_de_ambalaj + '.pdf',
                     str(self.corectie_1_folder_selected) + '/' + cor1 + '.pdf',
                     str(self.corectie_2_folder_selected) + '/' + cor2 + '.pdf']
            outputstr = factura_si_nir+'_'+nota_de_ambalaj+'_'+cor1+'_'+cor2
            for iterator in files:
                if iterator[-7:] == 'nan.pdf':
                    files.remove(iterator)
            try:
                self.merge_pdfs(paths=files, output=str(self.output_folder_selected) + '/' +
                                                    outputstr + '.pdf')
            except FileNotFoundError:
                pass
        messagebox.showinfo(title='PDFMerger', message='PDFs merged. You can find them here:\n' + str(self.output_folder_selected))


if __name__ == '__main__':
    test = PDF_merger_and_writer()