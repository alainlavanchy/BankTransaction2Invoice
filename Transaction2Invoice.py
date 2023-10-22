#!/usr/bin/env python3
"""
This script reads a pdf of a monthly account statement issued by the bank
comparing it with a csv list containing the qr code reference numbers and
the accordign invoice numbers.
If a match is found, the invoice number is printed next to the qr reference number.
"""
import pdfquery
import pandas as PD
import logging
import re
import os
import sys
import datetime
import yaml
import io
import xml.etree.ElementTree as ET
from PyPDF2 import PdfWriter, PdfReader, Transformation
from reportlab.pdfgen.canvas import Canvas
import tkinter as TK
#from tkinter import *
#from tkinter.ttk import *
#from tkinter import filedialog as fd
#from tkinter.messagebox import showinfo

class GenerateFromTemplate():
    """
    
    
    """
    def __init__(self,template):
        self.template_pdf = PdfReader(open(template, "rb"))
        self.template_page= self.template_pdf.pages[0]

        self.packet = io.BytesIO()
        self.c = Canvas(self.packet,pagesize=(self.template_page.mediabox.width,self.template_page.mediabox.height))

    
    def addText(self,text,point):
        self.c.drawString(point[0],point[1],text)

    def merge(self):
        self.c.save()
        self.packet.seek(0)
        result_pdf = PdfReader(self.packet)
        result = result_pdf.pages[0]

        self.output = PdfWriter()

        op = Transformation().rotate(0).translate(tx=0, ty=0)
        result.add_transformation(op)
        self.template_page.merge_page(result)
        self.output.add_page(self.template_page)
    
    def generate(self,dest):
        outputStream = open(dest,"wb")
        self.output.write(outputStream)
        outputStream.close()


def get_timestamp_user():
    user = os.getlogin()
    timestamp = datetime.datetime.now()
    txtstring = "{}: {}".format(timestamp, user)
    return txtstring


def read_configuration(config_path):
    """
    Parameters:
    config_path (path object): Path to the file

    The function read_configuration opens a yaml configuration file,
    extracts all data and returns the data as dictionary.
    """
    with open(config_path, "r") as yamlfile:
        config_data = yaml.load(yamlfile, Loader=yaml.FullLoader)
    return config_data

def check_file(filename):
    """
    Parameters:
    filename (str): Path to the file

    Returns:
    1 if file is found
    Sys.Exit(1) if an error occurs
    """
    if not isinstance(filename, str):
        logging.error('{} - Filename and path has to be a string.'.format(get_timestamp_user()))
        exit(1)
    if not os.path.isfile(filename):
        logging.error('{} - File not found'.format(get_timestamp_user()))
        exit(1)
    return 1

def create_xml_tree(filename):
    """
    Parameters:
    filename(str): Path to the PDF file

    Returns:
    Nothing, but creates the pdfXML.txt file.
    """
    pdf = pdfquery.PDFQuery(filename)
    pdf.load()
    pdf.tree.write('pdfXML.txt', pretty_print = True)

def check_arguments():
    """
    Parameters:
    None

    Returns:
    Arguments handed over in the comand line

    This function checks, if one argument is added to the command and gives it back.

    """
    logging.info("{} - Handed over arguments: {}".format(get_timestamp_user()), len(sys.argv))
    if len(sys.argv)==1:
        logging.critical("Missing arguments")
        exit(1)
    else:
        pdffile = sys.argv[1]
    return pdffile

def parseXML(xmlfile):
    """Parses through the xml file to find all the 27 digit QR reference numbers and their positions

    Args:
        xmlfile(str): Path to the xmlfile.

    Returns:
        payments(dict): Containing a list with the qr reference numbers and the locations within the pdf file.
    """

    #get the xml tag from the config file
    #xmltag = configdata['xml-tag']  remove # to be used in final script
    xmltag = "LTTextLineHorizontal"


    #create element tree object
    tree = ET.parse(xmlfile)

    #get root element
    root = tree.getroot()

    #create empty list for the payment elements
    payments = []

    tagfind = root.find(".//" + xmltag)
    if tagfind is None:
        logging.error("{} - XML Tag \" {} \" not valid".format(get_timestamp_user(), xmltag))
        exit(1)
    else:
    #iterate the pdf elements. Depending on the pdfs layout, the tag has to be changed
        for payment in root.iter(xmltag):
            string = str(payment.text)
            if re.match("\d{27}", string):
                paymentdetails = {}
                logging.debug("{} - Referencenumber: {}".format(get_timestamp_user(), payment.text))
                paymentdetails['Referencenumber'] = payment.text
                logging.debug("{} - Coordinates: {}".format(get_timestamp_user(), payment.get('bbox')))
                paymentdetails['Coordinates'] = payment.get('bbox')
                payments.append(paymentdetails)
        print(payments)
    return payments

def writeElements2pdf(payments):
    """Add the invoice number at the correct position to the pdf.

    Args:
        payments(list): List with all dictionaries with the QR reference numbers, invoice numbers and positions.

    Returns:
        newpdffile(str): Path to the new created pdf file with the added invoice numbers.    
    """
    if not isinstance(payments, list):
        logging.error('{} - var payments has to be a list.'.format(get_timestamp_user()))
        exit(1)
    try:
        pdf_file
    except NameError:
        logging.error('{} - pdf file has to be selected'.format(get_timestamp_user()))
        exit(1)
    newpdf = GenerateFromTemplate(pdf_file)
    movex = 10
    movey = 10
    for payment in payments:
        pos = payment["Coordinates"]
        xpos = re.search(r'\[\d*.\d{2}', pos)
        xpos = float(xpos[0][1:]) + movex
        ypos = re.search(r'\, \d*.\d{2}', pos)
        ypos = float(ypos[0][1:]) + movey
        print("x Position: {}, y Position {}".format(xpos, ypos))
        newpdf.addText(payment['Invoice ID'], (xpos, ypos))


def getinvoicenumbers(payments):
    """Gets the according invoice numbers from the csv file for each qr reference number found in the pdf.

    Args:
        payments(list): List with all the dictionaries containing the qr reference numbers and their positions within the pdf.

    Returns:
        payments(list): Same list with the dictionaries, added with the invoice number for each qr reference number.    
    """
    if not isinstance(payments, list):
        logging.error('{} - var payments has to be a list.'.format(get_timestamp_user()))
        exit(1)

    #Read the csv file
    try:
        csv_file
    except NameError:
        logging.error('{} - csv file has to be selected'.format(get_timestamp_user()))
        exit(1)

    #open CSV and create a data frame
    csv = PD.read_csv(csv_file)

    #add the invoice number to the dictionary using the information found in the csv
    for payment in payments:
        csvline = csv.loc[csv['QR Reference'] == payment['Referencenumber'].strip()]
        invoicenumber = (csvline['Invoice ID'].values[0])
        payment['Invoice ID'] = invoicenumber

    return payments

# Start visual elements    

def select_pdf():
    filetypes = (
        ('PDF Dateien', '*.pdf'),
        ('Alle Dateien', '*.*')
    )
    pdf_name = fd.askopenfilename(
        title='Datei oeffnen',
        initialdir=configdata['base-directory'] if os.path.isdir(configdata['base-directory']) else '/',
        filetypes=filetypes
    )

    pdf_name_name = pdf_name.strip()
    if (len(pdf_name)==0):
        showinfo("Fehler", "Es muss eine Datei ausgewählt werden")
        return
    else:
        pdf_file.set(pdf_name)

  

def select_csv():
    filetypes = (
        ('CSV Dateien', '*.csv'),
        ('Alle Dateien', '*.*')
    )
    csv_name = fd.askopenfilename(
        title='Datei oeffnen',
        initialdir=configdata['base-directory'] if os.path.isdir(configdata['base-directory']) else '/',
        filetypes=filetypes
    )

    csv_name = csv_name.strip()
    if (len(csv_name)==0):
        showinfo("Fehler", "Es muss eine Datei ausgewählt werden")
        return
    else:
        csv_file.set(csv_name)

def main():
    """
    This is the main function with all settings and commands to run the programm
    """
    config_file = "config.yml"
    global csv_file
    global pdf_file
    global configdata
    configdata = (read_configuration(config_file))
    logging.basicConfig(filename=configdata['log-filename'], level=logging.DEBUG)
    logging.info(configdata)
    logging.info('{} - Start Logging'.format(get_timestamp_user()))
    root = Tk()
    csv_file = StringVar()
    pdf_file = StringVar()
    root.title('Transaction 2 Invoice')
    root.resizable(True, True)
    root.geometry('900x500')
    p_icon = PhotoImage(file = 'Logo_icon.png')
    root.iconphoto(False, p_icon)
    logo_image = PhotoImage(file='Logo_fonts.png')
    canvas = Canvas(root, width=900, height=500)
    canvas.pack()
    canvas.create_image(20,20,anchor=NW, image=logo_image)
    canvas.configure(bg='white')
    lblPDFName  = Label(root, text = "PDF Datei Monatsauszug", width = 24)
    lblPDFName.place(x=50, y=100)
    txtPDFName  = Entry(root, textvariable = pdf_file, width = 80, font = ('bold 9'))
    txtPDFName.place(x=200, y=100)
    button_PDF = Button(
        root,
        text='PDF Monatsauszug',
        command=select_pdf
    )
    button_PDF.place(x=50, y=130)
    lblCSVName  = Label(root, text = "CSV Datei", width = 24)
    lblCSVName.place(x=50, y=180)
    txtCSVName  = Entry(root, textvariable = csv_file, width = 80, font = ('bold 9'))
    txtCSVName.place(x=200, y=180)
    button_CSV = Button(
        root,
        text='CSV Zuordnung',
        command=select_csv
    )
    button_CSV.place(x=50, y=210)

    button_run = Button(
        root,      
        text='Zuordnen',
        command=root.destroy
    )
    button_run.place(x=350, y=300)
    button_exit = Button(
        root,      
        text='Abbrechen',
        command=root.destroy
    )
    button_exit.place(x=450, y=300)

    root.mainloop()

    """

    pdffile = check_arguments()
    if check_file(pdffile) == 1:
        create_xml_tree(pdffile)
        logging.info('XML tree created')
    else:
        logging.error('PDF not able to open')
        exit(1)

    """

if __name__ == "__main__":
    #main()
    global csv_file
    global pdf_file
    pdf_file = "Kontoauszug_Januar.pdf"
    csv_file = "testdata.csv"
    logging.basicConfig(filename="xml.log", level=logging.DEBUG)
    payments = parseXML("pdfXML.txt")
    paymentswithinvoice = getinvoicenumbers(payments)
    writeElements2pdf(paymentswithinvoice)
