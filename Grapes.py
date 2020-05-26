#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# imports
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import pandas as pd
import requests
import numpy as np
import re
from bs4 import BeautifulSoup
import datetime

# set pandas settings
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# build GUI
class MyWindow:
    def __init__(self, win):
        
        self.lbl=Label(window, text = 'Running this program will:\n\n1. Retrieve grape listings from winebusiness.com\n\n2. Export listings to an excel spreadsheet titled "Grape_Data"\nwhich is saved on your desktop')
        self.lbl.place(relx=.5, rely=0.2, anchor=CENTER)
        
        self.b1=Button(win, text='Run program', command=self.run)
        self.b1.place(relx=0.5, rely=0.5, anchor=CENTER)

# main code
    def run(self):
        # create empty dataframe
        all_data = pd.DataFrame(index=[], columns=['Varietal','Type','Appellation','Qty','Price','Date','Listing_ID'])

        # find total # of records and set last page dynamically
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
        listings_found = str(max_result)
        print('Listings found = ' + listings_found)

        #Create loop through URL's to scrape
        for i in range(1,page_max,50):
            # Set URL
            URL = "https://www.winebusiness.com/classifieds/grapesbulkwine/?sort_type=1&sort_order=desc&start={}#anchor1".format(i)
            res = requests.get(URL)
            soup = BeautifulSoup(res.content,'lxml')

            # Define specific table
            table = soup.find("table", attrs={"class": "table wb-cl-table"})
            df = pd.read_html(str(table))[0]

            # Add Listing_ID's
            tbody = table.find("tbody")
            df['Listing_ID'] = [np.where(tag.has_attr('href'),tag.get('href'),"no link") for tag in tbody.find_all('a')]

            #Create output table
            all_data = pd.concat([all_data, df], ignore_index=True)

        listings_scraped = str(len(all_data))
        print('Listings scraped = ' + listings_scraped)

        if int(listings_found) - int(listings_scraped) == 0:
            print('No errors detected')
        else:
            print('ERROR FOUND: listings scraped does not equal listings found')

        #Add datestamp
        now = datetime.datetime.now()
        datestamp = now.strftime("%Y-%m-%d %H:%M:%S")
        all_data['Datestamp'] = datestamp

        # Create copy for analysis
        df = all_data.copy()

        #Filter on grapes only
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

        #Create pivot table1
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

        #Print pivot table1
        table1

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

        #print pivot table2
        table2

        #Export to Excel
        savefile = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                         ("All files", "*.*") ))
        writer = pd.ExcelWriter(f'{savefile}.xlsx', engine='xlsxwriter')
        table1.to_excel(writer, sheet_name="Summary")
        table2.to_excel(writer, sheet_name="CS by App")
        all_data.to_excel(writer, index=False, sheet_name="Listings")
        writer.save()
        
        # send success message and close
        messagebox.showinfo('Info', 'Process completed!')
        window.destroy()

window=Tk()
mywin=MyWindow(window)
window.title('Grape Scraper')
window.geometry("400x300+10+10")
window.mainloop()

