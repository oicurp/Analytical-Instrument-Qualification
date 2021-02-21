"""
Author: Renhuan Huang

Comment: 
the script is used to generate mass amount of word document based on a common template
the document name is constructed from reading parameters from csv file.

the next step is to auto populate the key words inside each word document

"""


from docx import Document
from csv import reader

# To open an existing document, use it as template

document=Document('(System ID), Requirements Traceability Matrix for (System Desc).docx')
parameter_file = 'Doc Parameter QC.csv'

# To read document id from a csv file, and store it into a list.

list_of_docid = []

with open(parameter_file,'r') as read_obj:
    csv_reader = reader(read_obj,)
    list_of_docid = list(csv_reader)

#print(list_of_docid[1][0])

for i in list_of_docid:
    doc_name = i[9] +', RTM for ' + i[0] + '.docx'
    print(doc_name)
    document.save(doc_name)
