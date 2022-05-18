from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd
import json

 
def addFont():
    pdfmetrics.registerFont(TTFont('Pangram', 'Font/Pangram.ttf'))
    pdfmetrics.registerFont(TTFont('Pangram-Black', 'Font/Pangram-Black.ttf'))
    pdfmetrics.registerFont(TTFont('Pangram-ExtraBold', 'Font/Pangram-ExtraBold.ttf'))
    pdfmetrics.registerFont(TTFont('Pangram-ExtraLight', 'Font/Pangram-ExtraLight.ttf'))
    pdfmetrics.registerFont(TTFont('Pangram-Medium', 'Font/Pangram-Medium.ttf'))
    pdfmetrics.registerFont(TTFont('Pangram-Bold', 'Font/Pangram-Bold.ttf'))

def addText():
    from PyPDF2 import PdfFileWriter, PdfFileReader
    import io
   

    type = ".pdf"
    excelDosyasi = (input("Veri alınacak excel dosyasının ismini giriniz (Örneğin '2022-rapor', uzantısını eklemeyiniz): "))
    excelDosyasi = "Excel Dosyaları/" + excelDosyasi+ ".xlsx"
    df = pd.read_excel(excelDosyasi)

    columnName = input("Yazıların alınacağı sütunun ismini giriniz: ")
    rangeValue = (input("Hangi satırlar arasını istiyorsunuz (Örneğin 20-30): "))
    x = int(input("Yazının yeri için x değerini giriniz (Dikey hizanın konumu): "))
    y = int(input("Yazının yeri için y değerini giriniz (Yatay hizanın konumu): "))
    rangeValue = rangeValue.split("-")
    min = int(rangeValue[0])
    max = int(rangeValue[-1])
    
    f = open('config.json')
    data = json.load(f)
    font = data["font"]
    fontSize = data["size"]

    in_pdf_file =  str(input("Yazı eklenecek pdf'in ismini giriniz (Örnek 'satislar', .pdf eklemeyiniz):  "))

    for i in range(min,max+1):
        out_pdf_file =  in_pdf_file +"-"+ str(i) + type
        text = df[columnName][i-2]
        packet = io.BytesIO()
        can = canvas.Canvas(packet)
        can.setFont(font, fontSize)        
        can.drawString(x, y, text)
        can.showPage()
        can.save()
        
        packet.seek(0)
    
        new_pdf = PdfFileReader(packet)
    
        existing_pdf = PdfFileReader(open("Değişecek PDF'ler/"+in_pdf_file+type, "rb"))
        output = PdfFileWriter()
    
        for i in range(len(existing_pdf.pages)):
            page = existing_pdf.getPage(i)
            page.mergePage(new_pdf.getPage(i))
            output.addPage(page)
    
        outputStream = open("Değişen PDF'ler/" +out_pdf_file, "wb")
        output.write(outputStream)
        outputStream.close()
        print(out_pdf_file + "Yazdırıldı")

 
 
addText()
print("İşlem Bitti..")
input("")