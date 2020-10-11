#! python3

import docx, os
from datetime import datetime

os.chdir(r'c:\Users\username\Documents\Reports') # "username" is a placeholder for your username

# Define "d" as document, load reference document from "refReport" folder within "Reports" folder
d = docx.Document('.\\refReport\\refReport.docx')

# Define current date for filename
todayNum = datetime.today().strftime('%Y-%m-%d')

# Save document
d.save(todayNum + ' Report.docx')

# Open the document in Word
os.startfile(todayNum + ' Report.docx')
