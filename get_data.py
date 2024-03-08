import os
from PyPDF2 import PdfReader
from openpyxl import Workbook
import re
from words2numsrus.extractor import NumberExtractor
import traceback
extractor = NumberExtractor()
# Create an Excel workbook
workbook = Workbook()
sheet = workbook.active

# Set the header row
header = ['Поставщик', 'Покупатель', 'Итого', 'date']
sheet.append(header)

# Directory containing your invoice files
directory = 'C:/Users/acer/OneDrive/Desktop/Dusel-Projects'

check = True
nameoffile = []
# Iterate over each file in the directory
for filename in os.listdir(directory):
    if filename.endswith('.pdf'):
        file_path = os.path.join(directory, filename)

        # Open the PDF file
        with open(file_path, 'rb') as file:
            try:
                pdf_reader = PdfReader(file)

                # Extract text from the first page
                first_page = pdf_reader.pages[0]
                text = first_page.extract_text()
                if len(pdf_reader.pages)>1:
                    text += pdf_reader.pages[1].extract_text()


                commission_agent_match = re.search(r'Поставщик:\s*(.*)', text)
                commission_agent = commission_agent_match.group(
                    1) if commission_agent_match else None

                # Find the "Покупатель" data
                buyer_match = re.search(r'Покупатель:\s*(.*)', text)
                buyer = buyer_match.group(1) if buyer_match else None
                if commission_agent is None:
                    commission_agent = re.search(r'Комиссионер:\s*(.*)', text)
                    commission_agent = commission_agent.group(1) if buyer_match else None

                # Find the "Итого" data
                total_match = re.search(r'Всего к оплате:\s*(.*)', text)
                total = total_match.group(1) if total_match else None

                match = re.search(r'оператор:[^\n]*\n([^оператор]*)', text)
                text_after_operator = match.group(1) if match else None

                # Extract the date from the text after оператор
                date_match = re.search(
                    r'\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}', text_after_operator)

                date = date_match.group(0).split(
                    ' ')[0] if date_match else None

                if buyer is not None and commission_agent is not None and total is not None and date_match is not None:
                    sheet.append([commission_agent, buyer,
                                  extractor.replace_groups(total), date])
                    file.close()

                    print("Поставщик:", commission_agent)
                    print("Покупатель:", buyer)
                    print("Итого:", extractor.replace_groups(total))
                    print("Date:", date)
                    print('-'*88)
                    os.remove(file_path)
                else:
                    file.close()
                
                sheet.column_dimensions['A'].width = 60
                sheet.column_dimensions['B'].width = 60
                sheet.column_dimensions['C'].width = 20
                sheet.column_dimensions['D'].width = 20
            except Exception as e:
                print(e)
                traceback.print_exc()


workbook.save('C:/Users/acer/OneDrive/Desktop/Dusel-Projects/file.xlsx')

