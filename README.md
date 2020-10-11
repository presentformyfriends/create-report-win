# createReportWin
Python script to create a "Weekly Report" DOCX document with current date and page numbers, saves with current date filename, and then opens the saved document with Microsoft Word.

Written for Windows.

## :memo: Usage

Copy or move the "Reports" folder to your Documents folder.

You can schedule this script in Task Scheduler.

![createReportWin.gif](img/createReportWin.gif)

## :snake: Dependencies

Written in Python for Windows. This script uses the following Python modules: docx, os, datetime.

Requires Microsoft Word, as the script creates a DOCX file. LibreOffice and/or OpenOffice should work too.