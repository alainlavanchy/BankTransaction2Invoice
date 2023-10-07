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

def main():
    """
    
    """
    logging.basicConfig(filename='Transaction2Invoice.log', level=logging.INFO)
    logging.info('Start Logging')
    pdffile = check_arguments()
    if check_file(pdffile) == 1:
        create_xml_tree(pdffile)
        logging.info('XML tree created')
    else:
        logging.error('PDF not able to open')
        exit(1)

if __name__ == "__main__":
    main()
    