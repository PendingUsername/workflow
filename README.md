# Workflow scripts written in Python. 
These are currently being updated, while some are ready to be released. 
------------------------------------------------------------------------------------------------
# 1. Exam ready
exam_ready.py is an email generator which generates an email notifying students that their exams are ready to download. It takes simple user input and generates an email template, saved as Exam_Notification.docx in the users documents folder.
# 2. Excel converter
excel_oasis.py allows the user to upload two excel files. File 1: the data extraction file containing the raw data of the excel file. File 2: The formatting file that defines the format into which File 1's data will be converted. A new excel file will be created, with the data from File 1 and the formatting from File 2. 
# 3. Grades ready
grades_ready is another email generator to notify users that their exam is ready to download. It also provides information about the exam length, location and time. Additional information about the exam is added to the email template.
# 4. Test examine
test_exam.py is compares two word documents, showing the difference between the two. This allows the user to quickly compare exams against one another. This make it easier to make edits within the examsoft platform. 