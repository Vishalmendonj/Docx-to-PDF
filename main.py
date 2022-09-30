import PyPDF2, docx, csv
from docx2pdf import convert
outputFile= open(f'C:/Users/Dinesh/PycharmProjects/adp project/csv files/users.csv','w',newline='')
outputWriter = csv.writer(outputFile)
outputWriter.writerow(['User Name','Phone','file(.docx)','No of para','file(.pdf)'])
user_name=input('Enter your Name: ')
user_phone=int(input('Enter your Phone number: '))
num_files= int(input("Enter the number of docx files to be converted to pdf files: "))
for num_file in range(num_files):
    user=input('Enter the .docx file to convert it into .pdf format: ')
    source_address=input('Please enter the source address: ')
    dest_address=input('Please enter the destination address: ')
    ask_encrypt=input("Do you want to encrypt the pdf file (y/n): ")

    doc= docx.Document(f'{source_address}/{user}.docx')
    doc_length=len(doc.paragraphs)
    if ask_encrypt.lower()=='n':
        converted_pdf=convert(f'{source_address}/{user}.docx',f'{dest_address}/{user}.pdf')

    elif ask_encrypt.lower()=='y':
        ask_password=input("Enter the password: ")
        converted_pdf = convert(f'{source_address}/{user}.docx',f'C:/Users/Dinesh/PycharmProjects/adp project/encrypted pdf/{user}.pdf')
        pdf_file= open(f'C:/Users/Dinesh/PycharmProjects/adp project/encrypted pdf/{user}.pdf','rb')
        pdfReader= PyPDF2.PdfFileReader(pdf_file)
        pdfWriter = PyPDF2.PdfFileWriter()
        for pagenum in range(pdfReader.numPages):
            pdfWriter.addPage(pdfReader.getPage(pagenum))

        pdfWriter.encrypt(ask_password)
        resultPdf = open(f'{dest_address}/{user}.pdf', 'wb')
        pdfWriter.write(resultPdf)
        resultPdf.close()

    else:
        print("Please answer with 'y' or 'n'")

    outputWriter.writerow([f'{user_name}',f'{user_phone}',f'{user}.docx',doc_length,f'{user}.pdf'])


