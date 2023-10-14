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
import os
import sys
import pathlib
import yaml
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

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
        logging.error('Filename and path has to be a string.')
        exit(1)
    if not os.path.isfile(filename):
        logging.error('File not found')
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
    logging.info("Handed over arguments: {}".format(len(sys.argv)))
    if len(sys.argv)==1:
        logging.critical("Missing arguments")
        exit(1)
    else:
        pdffile = sys.argv[1]
    return pdffile

def select_pdf():
    filetypes = (
        ('PDF Dateien', '*.pdf'),
        ('Alle Dateien', '*.*')
    )
    pdf_name = fd.askopenfilename(
        title='Datei oeffnen',
        initialdir='/',
        filetypes=filetypes
    )

    showinfo(
        title='Ausgewaehlte Datei',
        message=pdf_name
    )

def select_csv():
    filetypes = (
        ('CSV Dateien', '*.csv'),
        ('Alle Dateien', '*.*')
    )
    csv_name = fd.askopenfilename(
        title='Datei oeffnen',
        initialdir='/',
        filetypes=filetypes
    )

    showinfo(
        title='Ausgewaehlte Datei',
        message=csv_name
    )

def main():
    """
    This is the main function with all settings and commands to run the programm
    """
    config_file = "config.yml"
    configdata = read_configuration(config_file)
    logging.basicConfig(filename=configdata['log-filename'], level=logging.INFO)
    logging.info(configdata)
    logging.info('Start Logging')
    root = Tk()
    root.title('Transaction 2 Invoice')
    root.resizable(False, False)
    root.geometry('400x300')
    p_icon = PhotoImage(file = 'Logo_icon.png')
    root.iconphoto(False, p_icon)
    logo_image = PhotoImage(file='Logo_fonts.png')
    canvas = Canvas(root, width=400, height=300)
    canvas.pack()
    canvas.create_image(20,20,anchor=NW, image=logo_image)
    canvas.configure(bg='white')
    button_PDF = Button(
        root,
        text='PDF Monatsauszug',
        command=select_pdf
    )
    button_PDF.place(x=50, y=100)

    button_CSV = Button(
        root,
        text='CSV Zuordnung',
        command=select_csv
    )
    button_CSV.place(x=50, y=140)
    button_exit = Button(
        root,      
        text='Abbrechen',
        command=root.destroy
    )
    button_exit.place(x=50, y=200)

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
    main()
    