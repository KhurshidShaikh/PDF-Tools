import sys,os,pikepdf,io,pyttsx3
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QFileDialog,QInputDialog,QLabel,QVBoxLayout
from PyQt5.QtCore import QPropertyAnimation
from PyQt5.QtGui import QPixmap
from pdf2docx import Converter
from docx2pdf import convert
from PyPDF2 import PdfWriter,PdfReader
from reportlab.pdfgen import canvas


#  main application window
app = QApplication(sys.argv)

welcomeframe=QWidget()
pdftoolsframe = QWidget()

pdftools_layout = QVBoxLayout(pdftoolsframe)

welcome_label = QLabel("Welcome to PDF Tools", welcomeframe)
welcome_label.setGeometry(720, 300, 2000, 200)
welcome_label.setStyleSheet("font: bold italic 30pt 'Times New Roman'; color: #FFFFFF;")

pdf_icon_label = QLabel(welcomeframe)
base_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
image_path = os.path.join(base_dir, 'PDF.png')
pdf_icon_pixmap = QPixmap(image_path)  
pdf_icon_label.setPixmap(pdf_icon_pixmap)
pdf_icon_label.setGeometry(500, 300, 200, 200)

start_button = QPushButton("Start", welcomeframe)
start_button.setStyleSheet("border-radius:20px;font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#1061e3;")
start_button.setGeometry(700, 650, 500, 85)

welcomeframe.setGeometry(300, 300, 300, 200)
welcomeframe.setWindowTitle('PDF TOOLS')
welcomeframe.setStyleSheet("background-color: #0a0a0a; color: #FFFFFF;")

start_button.clicked.connect(pdftoolsframe.show)

button1 = QPushButton('📃Convert PDF To Word', pdftoolsframe)
button1.setGeometry(50, 50, 1800, 100)
button1.setStyleSheet("border-radius: 30px; font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button1)

button2 = QPushButton('📃Convert Word To PDF', pdftoolsframe)
button2.setGeometry(50, 200, 1800,100)
button2.setStyleSheet("font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button2)

button3 = QPushButton('🔑PDF Encryption', pdftoolsframe)
button3.setGeometry(50, 350, 1800, 100)
button3.setStyleSheet("font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button3)

button4 = QPushButton('📃+📃Merge PDFs', pdftoolsframe)
button4.setGeometry(50, 500, 1800, 100)
button4.setStyleSheet("font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button4)

button5 = QPushButton('📃PDF To text', pdftoolsframe)
button5.setGeometry(50, 650, 1800, 100)
button5.setStyleSheet("font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button5)

button6 = QPushButton('📃PDF Watermarking', pdftoolsframe)
button6.setGeometry(50, 770, 1800, 100)
button6.setStyleSheet("font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button6)

button7 = QPushButton('PDF To Audio🔉', pdftoolsframe)
button7.setGeometry(50, 900, 1800, 100)
button7.setStyleSheet("font: bold italic 20pt 'Times New Roman'; color: #FFFFFF;background-color:#2E2E2E;")
pdftools_layout.addWidget(button7)


pdftoolsframe.setStyleSheet("background-color: #0a0a0a; color: #FFFFFF;")




def pdf_to_word():
    pdf_file,_ = QFileDialog.getOpenFileName(pdftoolsframe, 'Select file', 'C:\\',filter='PDF Files (*.pdf)')
    if not pdf_file:
        return 
   
    word_file = os.path.join(os.path.expanduser("~"), "Downloads", f"{os.path.splitext(os.path.basename(pdf_file))[0]}.docx")
     
    
    try:
        cv = Converter(pdf_file)
        cv.convert(word_file)
        cv.close()
        QMessageBox.information(pdftoolsframe,"Message","converted to word successfully")
        QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")
    except Exception as e:
        print(e)

    
   
def word_to_pdf():
    word_file,_=QFileDialog.getOpenFileName(pdftoolsframe,'Select File','C:\\',filter='Word Files (*.docx)')
    if not word_file:
        return
    # word_file=word_file[0]
    pdf_file=os.path.join(os.path.expanduser('~'),'Downloads',f"{os.path.splitext(os.path.basename(word_file))[0]}.pdf")
    
    try:
         convert(word_file,pdf_file)
         QMessageBox.information(pdftoolsframe,"Message","converted to PDF successfully")
         QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")
    except Exception as e:
         print(e)



def pdf_encryption():
    pdf_file,_=QFileDialog.getOpenFileName(pdftoolsframe,"Select File",'C:\\',filter='PDF Files (*.pdf)')
    if not pdf_file:
        return
    password,ok=QInputDialog.getText(pdftoolsframe,"Input","Set a Password")
    open_pdf=pikepdf.Pdf.open(pdf_file)
    no_extr=pikepdf.Permissions(extract=False)
    protected_file = os.path.join(os.path.expanduser("~"), "Downloads", f"{os.path.splitext(os.path.basename(pdf_file))[0]}_protected.pdf")
    open_pdf.save(protected_file,encryption=pikepdf.Encryption(user=password,allow=no_extr))
    QMessageBox.information(pdftoolsframe,'Message','PDF Encrypted Successfully')
    QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")


def pdf_merger():
    pdf_file,_=QFileDialog.getOpenFileNames(pdftoolsframe,"Select Files","C:\\",filter='PDF files (*.pdf)')
    if not pdf_file:
        return
    selected_files=[]
    for pdf_file in pdf_file:
        selected_files.append(os.path.abspath(pdf_file))
    print(selected_files)
    
    merger=PdfWriter()
    for pdf in selected_files:
        merger.append(pdf)
    
    merged_file = os.path.join(os.path.expanduser("~"), "Downloads", f"{os.path.splitext(os.path.basename(pdf_file))[0]}_merged.pdf")
    merger.write(merged_file)
    merger.close()
    QMessageBox.information(pdftoolsframe,"Message","PDFs Merged Successfully")
    QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")

    

def pdf_to_text():
    pdf_file,_=QFileDialog.getOpenFileName(pdftoolsframe,'Select File','C:\\',filter='PDF Files (*.pdf)')
    if not pdf_file:
        return
    try:
       file=PdfReader(pdf_file)
       text=""
       for page in file.pages:
           text+=page.extract_text()
       text_file = os.path.join(os.path.expanduser("~"), "Downloads", f"{os.path.splitext(os.path.basename(pdf_file))[0]}.txt")
       with open(text_file,'w',encoding='utf=8') as txtfile:
          
            txtfile.write(text)
       QMessageBox.information(pdftoolsframe,'Message',"Text File Generated")
       QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")
    except Exception as e:
        QMessageBox.critical(pdftoolsframe, 'Error', f'Error converting PDF to audio: {str(e)}')
    

def pdf_watermark():
    pdf_file,_=QFileDialog.getOpenFileName(pdftoolsframe,"Select File","C:\\",filter="PDF (*.pdf)")
    if not pdf_file:
        return
    watermark_text,ok=QInputDialog.getText(pdftoolsframe,'Input','Enter Your WaterMark Text')
    
    output_pdf=os.path.join(os.path.expanduser("~"), "Downloads", f"{os.path.splitext(os.path.basename(pdf_file))[0]}_watermarked.pdf")
    transparency=0.5
    try:
        with open(pdf_file, 'rb') as file:
                pdf_reader =PdfReader(file)
                pdf_writer=PdfWriter(file)

        #  canvas for adding text
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(pdf_reader.pages[0].mediabox.width, pdf_reader.pages[0].mediabox.height))
                can.setFont("Helvetica", 80) 
                can.setFillAlpha(transparency)  # Set transparency
                can.setFillColorRGB(0, 0, 0)  # Set text color (black)
                can.drawCentredString(float(pdf_reader.pages[0].mediabox.width/ 2),float(pdf_reader.pages[0].mediabox.height/ 2), str(watermark_text))  # Center the text diagonally
                can.save()

        # Move the canvas to the beginning of the StringIO buffer
                packet.seek(0)
                text = PdfReader(packet)

        # Iterate through each page in the input PDF
                for page_number in range(len(pdf_reader.pages)):
            # Get the page
                    page = pdf_reader.pages[page_number]

            # Merge the watermark with the page
                    page.merge_page(text.pages[0])

            # Add the modified page to the PdfFileWriter
                    pdf_writer.add_page(page)

        # Saving the output PDF file with the watermark
                with open(output_pdf, 'wb') as output_file:
                    pdf_writer.write(output_file)
        QMessageBox.information(pdftoolsframe,'Message','PDF Watermarked Succesfully')
        QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")
    except Exception as e:
        QMessageBox.critical(pdftoolsframe, 'Error', f'Error : {str(e)}')



def pdf_to_audio():
    pdf_file,_=QFileDialog.getOpenFileName(pdftoolsframe,'Select File','C:\\',filter='PDF File (*.pdf)')
    if not pdf_file:
        return
    # text-to-speech engine
    engine = pyttsx3.init()
    try:
        with open(pdf_file, 'rb') as file:
            pdf_reader = PdfReader(file)

            # Extract text from each page
            text_content = ''
            for page_number in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_number]
                text_content += page.extract_text()

            audio_file =os.path.join(os.path.expanduser("~"), "Downloads", f"{os.path.splitext(os.path.basename(pdf_file))[0]}.mp3") 
            engine.save_to_file(text_content,audio_file)

            # Wait for the speech to be generated
            engine.runAndWait()

            QMessageBox.information(pdftoolsframe, 'Message', 'PDF Converted to Audio Successfully')
            QMessageBox.information(pdftoolsframe,"Message","File Saved in Downloads")
            

    except Exception as e:
        QMessageBox.critical(pdftoolsframe, 'Error', f'Error converting PDF to audio: {str(e)}')
        




#  button click event to the action
button1.clicked.connect(pdf_to_word)
button2.clicked.connect(word_to_pdf)
button3.clicked.connect(pdf_encryption)
button4.clicked.connect(pdf_merger)
button5.clicked.connect(pdf_to_text)
button6.clicked.connect(pdf_watermark)
button7.clicked.connect(pdf_to_audio)


# Set the window properties
pdftoolsframe.setGeometry(300, 300, 300, 200)
pdftoolsframe.setWindowTitle('PDF TOOLS')

pdftoolsframe.setWindowOpacity(0.0)
animation = QPropertyAnimation(pdftoolsframe, b"windowOpacity")
animation.setDuration(1000)
animation.setStartValue(0.0)
animation.setEndValue(1.0)
animation.start()

# Show the window
welcomeframe.show()
pdftoolsframe.hide()
# Start the application event loop
sys.exit(app.exec_())



