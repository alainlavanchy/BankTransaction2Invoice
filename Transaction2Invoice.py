#!/usr/bin/env python3
"""
This script reads a pdf of a monthly account statement issued by the bank
comparing it with a csv list containing the qr code reference numbers and
the accordign invoice numbers.
If a match is found, the invoice number is printed next to the qr reference number.
"""
import pdfquery
import pandas as pd
import logging
import re
import os
import sys
import datetime
import yaml
import xml.etree.ElementTree as ET
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

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
    """ 
    Parameters:
    xmlfile(str): Path to the xmlfile

    Returns:
    payments(dict): Containing a list with the qr reference numbers and the locations within the pdf file.

    This function search for all the 27 digit qr refrence numbers within the xml file and saves
    the found numbers aswell as the allocated coordinates in the payments dictionary.
    
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
                logging.debug("{} - Referenznummer: {}".format(get_timestamp_user(), payment.text))
                paymentdetails['Referenznummer'] = payment.text
                logging.debug("{} - Koordianten: {}".format(get_timestamp_user(), payment.get('bbox')))
                paymentdetails['Koordinaten'] = payment.get('bbox')
                payments.append(paymentdetails)
        print(payments)
    return payments

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
    logging.basicConfig(filename="xml.log", level=logging.DEBUG)
    parseXML("pdfXML.txt")