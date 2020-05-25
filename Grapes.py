#!/usr/bin/env python
# coding: utf-8

# In[1]:


# imports
import pandas as pd
import requests
import numpy as np
import re
from bs4 import BeautifulSoup
import datetime

# pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# create empty dataframe
all_data = pd.DataFrame(index=[], columns=['Varietal','Type','Appellation','Qty','Price','Date','Listing_ID'])
all_data


# In[2]:


# find final record and set maximum page
URL = "https://www.winebusiness.com/classifieds/grapesbulkwine/?sort_type=1&sort_order=desc&start=1#anchor1"
res = requests.get(URL)
soup = BeautifulSoup(res.content,'lxml')
searched_word = 'Results'
find_string = soup.body.find(text=re.compile(searched_word), recursive=True)
def largestNumber(in_str):
    l=[int(x) for x in in_str.split() if x.isdigit()]
    return max(l) if l else None
max_result = largestNumber(find_string)
page_max = ((int(max_result/50))*50)+2

# check max result to verify correct number of records
print('Max record is ' + str(max_result))
print('Max page is ' + str(page_max))


# In[3]:


#Create loop through URL's
for i in range(1,page_max,50):
    # Set URL
    URL = "https://www.winebusiness.com/classifieds/grapesbulkwine/?sort_type=1&sort_order=desc&start={}#anchor1".format(i)
    res = requests.get(URL)
    soup = BeautifulSoup(res.content,'lxml')
    print(URL)

    # Define specific table
    table = soup.find("table", attrs={"class": "table wb-cl-table"})
    df = pd.read_html(str(table))[0]

    # Add Listing_ID's
    tbody = table.find("tbody")
    df['Listing_ID'] = [np.where(tag.has_attr('href'),tag.get('href'),"no link") for tag in tbody.find_all('a')]
    
    #Create output table
    all_data = pd.concat([all_data, df], ignore_index=True)
    
print('# of records scraped = ' + str(len(all_data)))


# In[4]:


# view final output dataframe
all_data


# In[5]:


#Add datestamp
now = datetime.datetime.now()
datestamp = now.strftime("%Y-%m-%d %H:%M:%S")
all_data['Datestamp'] = datestamp
all_data

# In[6]:

# Create copy for analysis
df = all_data.copy()

#Filter on grapes only
    #is_grapes =  df['Type']=='Grapes'
    #df['Grapes'] = is_grapes
df = df[(df['Type']=='Grapes')]

#Create State column and filter on California
df['State'] = df['Appellation'].str[:2]
df = df[(df['State']=='CA')]

#Remove state from appellation values
df['Appellation'] = df['Appellation'].str[5:]

#Filter certain varietals
df = df[df['Varietal'].str.contains("Cabernet Sauvignon|Merlot|Pinot Noir|Chardonnay|Sauvignon Blanc", case=False)]
df = df[~df['Varietal'].str.contains("Sold", case=False)]

#Remove nulls from price and quantity and cast as float
df = df[df['Qty'].notna()]
NewQty = df['Qty'].str.extract('(\d*\.\d+|\d+)', expand=False).astype(float)
df = df[df['Price'].notna()]
NewPrice = df['Price'].str.extract('(\d*\.\d+|\d+)', expand=False).astype(float)
df['Tons'] = NewQty
df['$/Ton'] = NewPrice

#Add total cost
TotalCost = NewQty * NewPrice
df['Total Cost'] = TotalCost
df


# In[7]:


table1 = pd.pivot_table(df, values=['Tons', 'Total Cost'],
                     index=['Varietal'],aggfunc=np.sum, margins=False)

table1['$/Ton'] = table1['Total Cost'] / table1['Tons']

#Format Numbers
table1['Tons'] = table1['Tons'].map('{:,.1f}'.format)
table1['Total Cost'] = table1['Total Cost'].map('${:,.0f}'.format)
table1['$/Ton'] = table1['$/Ton'].map('${:,.0f}'.format)

#Sort
table1 = table1.sort_values(by=['$/Ton'], ascending=False)

#Reorder Columns
table1.columns = ['Tons Available', 'Total Value', 'Avg $/Ton']

#Print
table1


# In[8]:


#Create pivot table2
table2 = pd.pivot_table(df[df.Varietal == "'20 Cabernet Sauvignon"], values=['Tons', 'Total Cost'],
                     index=['Appellation'],aggfunc=np.sum, margins=False)

#Create avg price per ton column
table2['$/Ton'] = table2['Total Cost'] / table2['Tons']

#Sort by avg price
table2 = table2.sort_values(by=['$/Ton'], ascending=False)

#Format numbers
table2['Tons'] = table2['Tons'].map('{:,.1f}'.format)
table2['Total Cost'] = table2['Total Cost'].map('${:,.0f}'.format)
table2['$/Ton'] = table2['$/Ton'].map('${:,.0f}'.format)

#print
table2


# In[9]:


#Export to Excel
import os
username = os.getlogin()
#all_data.to_excel(f'C:\\Users\\{username}\\Desktop\\grape_data.xlsx', index = False)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(f'C:\\Users\\{username}\\Desktop\\Grape_Data.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
table1.to_excel(writer, sheet_name='Summary')
table2.to_excel(writer, sheet_name='Cab Sauv by App')
all_data.to_excel(writer, sheet_name='Raw Data')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

