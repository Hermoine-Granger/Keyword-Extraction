import spacy
import pytextrank
import PyPDF2
from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import os
from docx2pdf import convert


#Function to implement TextRank
def TextRank(text):

    nlp = spacy.load("en_core_web_sm")
    # add PyTextRank to  spaCy
    tr = pytextrank.TextRank()
    nlp.add_pipe(tr.PipelineComponent, name="textrank", last=True)

    doc = nlp(text)


    f = open("KeyWords.txt", "w")
    f = open("KeyWords.txt", "a")
    for p in doc._.phrases:

        #print(p.text)
        f.write(p.text+'\n')

    f.close()


#This function reads data from text file and passes it in TextRank function
def TextFileRead(input_path):
    fileObject = open(input_path, "r")
    data = fileObject.read()

    TextRank(data)


#This function converts PDF to text
def pdfToImageToText(PDF_file):

    #1 Converting PDF to images


    # We will store all pages of PDF in a variable
    pages = convert_from_path(PDF_file, 500, poppler_path=r"C:\Poppler\Release-20.11.0\poppler-20.11.0\bin")

    # Set Ccunter to store each page of PDF to image
    image_counter = 1

    # Save all the images of pages stored above
    for page in pages:
        # Give name to each file for each page of PDF as JPG
        # PDF page n -> page_n.jpg
        filename = "page_" + str(image_counter) + ".jpg"

        # Save the image file in system
        page.save(filename, 'JPEG')

        # Increase counter update filename
        image_counter = image_counter + 1


    #2 - Use OCR to Read text from the imagefiles


    # Keep record of total number of pages
    filelimit = image_counter - 1

    # Text file creation to store the output
    outfile = "out_text.txt"

    # Use the append mode to add data of all images to the same file
    open(outfile, 'w').close()
    f = open(outfile, "a")

    # Iterating all pages
    for i in range(1, filelimit + 1):
        # Set filename to recognize text from
        # page_n.jpg
        filename = "page_" + str(i) + ".jpg"

        # Use Pytesserct to extract text as string from image
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        text = str(((pytesseract.image_to_string(Image.open(filename)))))

        # The recognized text is stored in variable text
        # We will replace every '-\n' to '' because many times
        # in PDFs, if a word can't be written fully  in line ending, a 'hyphen' is added.
        # And the rest of the word is written in the next line
        text = text.replace('-\n', '')

        # Write the text after above preprocessing to the file.
        f.write(text)

    # Closing the file.
    f.close()

    #delete the saved images stored in your folder
    for i in range(1, filelimit + 1):
        # Set filename to delete
        # page_n.jpg
        filename = "page_" + str(i) + ".jpg"
        os.remove(filename)


    return outfile


#This function returns path of created pdf after converting docx to pdf

def DocxToPDF(infile):
    convert(infile)
    outfile='SampleDocx.pdf'
    convert(infile,outfile)
    return outfile

if __name__ == "__main__":
    print ("Enter \n 1 for pdf \n 2 for docx \n")
    input_type=int(input())
    if input_type not in (1,2,3):
        print ("Invalid input")


    if input_type==1:
        print("Give location to pdf file:\n")
        input_path = input()
        # Use below function to read data from above pdf
        text_file=pdfToImageToText(input_path)
        # Apply TextRank algorithm for Keyword extraction
        TextFileRead(text_file)
        print("Extracted Keywords are saved in the same folder in KeyWords.txt ")


    elif input_type==2:
        print("Give location to docx file:\n")
        input_path = input()

        #create temporary pdf file from docx
        outfile=DocxToPDF(input_path)
        #Use below function to read data from above pdf
        text_file = pdfToImageToText(outfile)
        #Apply TextRank algorithm for Keyword extraction
        TextFileRead(text_file)
        #remove temporary pdf file
        os.remove(outfile)
        print("Extracted Keywords are saved in the same folder in KeyWords.txt ")











