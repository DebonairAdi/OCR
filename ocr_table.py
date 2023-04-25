# import dependencies
import re, PIL, pytesseract, logging
import pandas as pd
from os import fsdecode, listdir, path
from img2table.ocr import TesseractOCR
from img2table.document import Image
from pdf2image import convert_from_path

# path of the input directory
input_path = "input"

# path of the output directory
output_path = "output"

# poppler exe path for PDF related processes
path_to_poppler_exe = r"poppler-0.68.0\bin"

# an empty dataframe
dataframe = pd.DataFrame(list())

# writing empty DataFrame to the new excel file
dataframe.to_excel('output.xlsx',index= False,)

# create log file
def create_logging():
    
    """
    Create a log file
    ---------------------------
    This function is used to create
    and configure the logging for the
    log file that stores the log messages.
    """
    
    # set logging configuration
    logging.basicConfig(filename="log_file.log", filemode='a', format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s', datefmt='%H:%M:%S', level=logging.INFO)
    return logging

# get invoice number from text
def get_invoice_no(text : str):
    
    """
    Get the Invoice Number
    -----------------------------
    This method is used to get the
    invoice number from the text string.
    It uses regular expressions to
    extract the invoice number from
    the text.

    Parameters
    ----------
    text : str
        It is the text of the image
        from which the invoice number
        is to be extracted.
    """
    
    # possible regex patterns
    regex_list = ["(?<=Invoice Number:).*","(?<=INVOICE).*"]
    
    # iterate over the regex
    for regex in regex_list:
        
        # look for invoice number
        result = re.findall(regex, text)
        
        # if result is found
        if len(result)>0:
            break
        
    return str(result[0]).lstrip().rstrip()

# get invoice date from text
def get_invoice_date(text : str):
    
    """
    Get the Invoice Date
    -----------------------------
    This method is used to get the
    invoice date from the text string.
    It uses regular expressions to
    extract the invoice date from
    the text.

    Parameters
    ----------
    text : str
        It is the text of the image
        from which the invoice date
        is to be extracted.
    """
    
    # possible regex patterns
    regex_list = ["(?<=Invoice Date:).*","(?s)(?<=\n\nDATE)(.*?)(?=PLEASE)"]
    
    # iterate over the regex list
    for regex in regex_list:
        
        # look for the invoice date
        result = re.findall(regex, text)
        
        # if invoice date is found
        if len(result)>0:
            break
        
    return str(result[0]).lstrip().rstrip()
    
# get invoice address from text
def get_invoice_address(text : str):
    
    """
    Get the Invoice Address
    -----------------------------
    This method is used to get the
    invoice address from the text string.
    It uses regular expressions to
    extract the invoice address from
    the text.

    Parameters
    ----------
    text : str
        It is the text of the image
        from which the invoice address
        is to be extracted.
    """
    
    # default value set to empty string
    address = ""
    
    # possible regex patterns
    regex_list = ["(?s)(?<=Shipped To[)]:)(.*?)(?=# Description)","(?s)(?<=BILL TO)(.*?)(?=SHIP DATE)"]
    
    # iterate over regex list
    for regex in regex_list:
        
        # look for the address
        result = re.findall(regex, text)
        
        # if address is found
        if len(result)>0:
            break
        
    # raw extracted address
    add = str(result[0]).lstrip().rstrip()
    
    # remove duplicate words from the address
    for line in add.splitlines():
        non_dup = sorted(set(str(line).split()), key=str(line).split().index)
        address = address + " " +" ".join(non_dup)
    
    return str(address).lstrip().rstrip()

# get invoice subtotal
def get_invoice_subtotal(text : str):
    
    """
    Get the Invoice Subtotal
    -----------------------------
    This method is used to get the
    invoice subtotal from the text string.
    It uses regular expressions to
    extract the invoice subtotal from
    the text.

    Parameters
    ----------
    text : str
        It is the text of the image
        from which the invoice subtotal
        is to be extracted.
    """
    
    # look for the subtotal
    try:
        
        # extract the subtotal
        result = re.findall("(?<=Subtotal:).*", text)
        return [str(x).lstrip().rstrip() for x in result]
    except:
        
        # if subtotal not found
        result = None
        return result

# get invoice total
def get_invoice_total(text : str):
    
    """
    Get the Invoice total
    -----------------------------
    This method is used to get the
    invoice total from the text string.
    It uses regular expressions to
    extract the invoice total from
    the text.

    Parameters
    ----------
    text : str
        It is the text of the image
        from which the invoice total
        is to be extracted.
    """
    
    # split the text into lines
    splited_text = text.splitlines()
    
    # extract the invoice total
    try:
        # look for invoice total
        req_line = "".join([line for line in splited_text if str(line).__contains__("Total Amounts (INR)")])
        result = str(req_line).split()[-1]
        return str(result).lstrip().rstrip()
    
    # if invoice total not found
    except:
        result = None
        return result

# get the table from an image
def get_table_From_image(input_file : path, output_file : path):
    
    """
    Extract the table from an image
    -----------------------------
    This method is used to extract 
    the table from an image using the
    OCR and Computer Vision algorithms

    Parameters
    ----------
    input_file : path
        It is the path of image
        which is to be processed.
    output_file : path
        It is the path where the 
        output will be saved.
    """
    
    # Instantiation of OCR
    ocr = TesseractOCR(n_threads=1, lang="eng")

    # Instantiation of an image
    img = Image(input_file)
    
    # Extraction of tables and creation of an xlsx file containing tables
    img.to_xlsx(dest=output_file,
                ocr=ocr,
                implicit_rows=False,
                borderless_tables=True,
                min_confidence=50)

# get dataframe from excels to process with custom transformation
def get_df(excel_path : path, invoice_no : str, invoice_date : str, invoice_address : str, invoice_total : str, invoice_subtotal : list):
    
    """
    Extract the dataframe from an excel
    -----------------------------
    This method is used to extract 
    the dataframe from an excel in order
    to perform custom transformations

    Parameters
    ----------
    excel_path : path
        It is the path of the
        excel file.
    invoice_no : str
        It is the invoice number
        from the invoice.
    invoice_date : str
        It is the invoice date
        from the invoice.
    invoice_address : str
        It is the invoice address
        from the invoice.
    invoice_total : str
        It is the invoice total
        from the invoice.
    invoice_subtotal : str
        It is the invoice subtotal
        from the invoice.
    """
    
    # read the excel file and convert it into a dataframe
    df = pd.read_excel(excel_path, "Page 1 - Table 2")
    
    # add new columns
    df['Invoice No.'] = invoice_no
    df['Invoice Date'] = invoice_date
    df['Invoice Address'] = invoice_address
    if invoice_total is not None:
        df['Invoice Total'] = invoice_total
    if len(invoice_subtotal) >0:
        df['Invoice Subtotal'] = ", ".join(invoice_subtotal)
        
    # add $ symbol to the amount field
    try:
        df['AMOUNT'] = ["$"+str(x) if "." in str(x) else x for x in df["AMOUNT"] ]
    except:
        pass
    
    return df

# main calling function
def main_caller():
    
    """
    Main calling function
    -----------------------------
    It is the main calling function
    which has all the logic implemented.
    """
    
    logger = create_logging()
    
    logger.info("Process started...")
    
    # iterate through the input files for PDF
    for file in listdir(input_path):
        
        logger.info("Looking for PDF files in the directory")
        
        # get the filename
        filename = fsdecode(file)
        
        # check if the file is a PDF
        if filename.endswith("pdf"):
        
            # complete input file path
            input_file = path.join(input_path, filename)    
            
            # convert PDF into images
            images = convert_from_path(input_file , poppler_path = path_to_poppler_exe)
            
            # iterate through the images
            for i in range(len(images)):
            
                # Save pages as images in the pdf
                images[i].save(path.join(input_path, str(filename).replace(".pdf", ".png")) , 'PNG')
                
    logger.info("PDF converted into images.")
    
    # iterate through the input directory to extract tables and data from images
    for idx, file in enumerate(listdir(input_path)):
        
        # get the filename
        filename = fsdecode(file)
        
        # check if the file is an image
        if filename.endswith("png"):
            
            # complete input file path
            input_file = path.join(input_path, filename)
            
            # complete output file path
            output_file = path.join(output_path, f"img_table_{idx}.xlsx")
            
            # create the excels with extracted tables from images
            get_table_From_image(input_file, output_file)
            
            logger.info("Tables extracted from image")
            
            # get complete text from the image using OCR
            text = str(((pytesseract.image_to_string(PIL.Image.open(input_file)))))

            # join hyphen ending newlines in PDF if any
            text = text.replace("-\n", "")
            
            # get the invoice number
            invoice_no = get_invoice_no(text)
            
            # get the invoice date
            invoice_date = get_invoice_date(text)
            
            # get the invoice address
            address = get_invoice_address(text)
            
            # get the invoice total
            total = get_invoice_total(text)
            
            # get the invoice subtotal
            subtotal = get_invoice_subtotal(text)
            
            logger.info("Other metadata fields extracted from the image.")
            
            # create dataframe from the excel file generated containing the tables
            df = get_df(output_file, invoice_no, invoice_date, address, total, subtotal)
            
            logger.info("Dataframe created to store")
            
            # store results to excel file
            with pd.ExcelWriter('output.xlsx', engine='openpyxl', mode='a') as writer:  
                df.to_excel(writer, sheet_name= f'Results_{idx}', index= False,)
    
    logger.info("Final Output file generated successfully")
                
                
if __name__ == '__main__':
    
    # call the main function
    main_caller()

    
            