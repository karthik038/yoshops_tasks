#!/usr/bin/env python
# coding: utf-8

# #### Python program to Create a excel file

# In[1]:


#pip install xlsrwriter

import xlsxwriter

workbook = xlsxwriter.Workbook('Score_card.xlsx')
worksheet = workbook.add_worksheet("Mark sheet")
scores = (
    ['Math', 93],
    ['Science',   91],
    ['Social',  87],
    ['Tamil',    83],
    ['English', 83]
)

row = 0
col = 0
 
for name, score in (scores):
    worksheet.write(row, col, name)
    worksheet.write(row, col + 1, score)
    row += 1
 
workbook.close()


# #### Python program for Import data from an excel file 

# In[59]:


#pip install pandas

import pandas as pd

df = pd.read_excel("Score_card.xlsx")
print(df)


# #### Python program for Format data in excel sheet

# In[1]:


import xlsxwriter
workbook = xlsxwriter.Workbook('company_employee_data.xlsx')
worksheet = workbook.add_worksheet()
 
cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format.set_font_size(16)
cell_format.set_underline(2)
cell_format.set_align('center')
 
cell_format1 = workbook.add_format({'font_color': 'blue'})
 
cell_format1.set_align('center')
worksheet.write('A1', 'Name', cell_format)
worksheet.write('B1', 'Department', cell_format)
row = 1
col = 0
 
data = ( 
    ['Rajendra', 'Business Analyst'], 
    ['Kashish','Voice and Non_voice Process '], 
    ['Arun', 'Data Associate'], 
    ['Rohan','Cloud Associate'], 
) 
 
worksheet.set_column('B1:B1', 60)
worksheet.set_column('B2:B5',60,cell_format1)
worksheet.set_column('A1:A5', 20,cell_format1)
 
for name, score in (data): 
    worksheet.write(row, col, name) 
    worksheet.write(row, col + 1, score) 
    row += 1
workbook.close()


# #### Python program for Prepare excel charts Pie Chart and Bar Chart 

# In[13]:


workbook = xlsxwriter.Workbook('chart_pie_bar.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})

headings = ['Category', 'Values']
data = [
   ['Apple', 'Cherry', 'Pecan'],
   [60, 30, 10],
] 

worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])

#Pie chart

chart = workbook.add_chart({'type': 'pie'})
chart.add_series({
   'name':       ['Sheet1', 0, 2],
   'categories': ['Sheet1', 1, 0, 3, 0],
   'values':     ['Sheet1', 1, 1, 3, 1],
})

chart.set_title({'name':'Pie chart for sales data'})
chart.set_style(10)
worksheet.insert_chart('C2', chart, {'x_offset': 25, 'y_offset': 10})


#Bar chart

chart1 = workbook.add_chart({'type': 'bar'})
chart1.add_series({
   'name':       ['Sheet1', 0, 2],
   'categories': ['Sheet1', 1, 0, 3, 0],
   'values':     ['Sheet1', 1, 1, 3, 1],
})

chart1.set_title({'name':'bar chart for sales data'})
chart1.set_style(10)
worksheet.insert_chart('K2', chart1)

workbook.close()


# #### Python program for Extract mobile no from PDF, XML and MS word file and save into MS excel

# In[ ]:


#PDF


# In[7]:


#pip install PyPDF2

import PyPDF2  
import re
import pandas as pd

pdfFileObj = open('customer_contact_data.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

pageObj = pdfReader.getPage(0)
content =  pageObj.extractText()
# print(content)
# print("---------------------------------------------------")

number = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]',content)
print(number)


# In[55]:


#creating a csv file
import csv

Pdf_df = pd.DataFrame(list(zip(number)), columns=['Mobile_No'])
Pdf_df.to_csv("PdfToExcel.csv")


# In[ ]:


#Ms Word


# In[6]:


#pip install docx2txt

import docx2txt
result = docx2txt.process("Customer_contact_data.docx")
print(result)


# In[7]:


numbers = re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]',result)
print(numbers)

#creating a csv file
word_df = pd.DataFrame(list(zip(numbers)), columns=['Mobile_No'])
word_df.to_csv("WordToExcel.csv")


# In[10]:





# In[ ]:




