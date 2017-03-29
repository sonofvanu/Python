import urllib.request
import re
import sqlite3
import xlsxwriter
from pyexcel_xls import read_data
from pyexcel_xls import save_data
from bs4 import BeautifulSoup

#Created an empty carrier or container for the objects from the excel sheet
URL = ''
keyword1=keyword2=keyword3=keyword4=''

#Created a function in order to fetch the url using the urllib library
def seo_fetch_page(siteURL):
    site= siteURL

    hdr = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
           'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
           'Accept-Encoding': 'none',
           'Accept-Language': 'en-US,en;q=0.8',
           'Connection': 'keep-alive'}

    req = urllib.request.Request(site, headers=hdr)

    try:
        page = urllib.request.urlopen(req)
    except urllib.request.HTTPError as e:
        print (e.fp.read())

    content = page.read().decode('utf-8')
    return content

#A function in order to create a table in the database for the data obtained
def seo_createTable():
    #Open the existing database
    conn = sqlite3.connect('E:\\vijay\\seo_program\\seo\\program\\Project\\Project.db')
    print ("Opened database successfully");

    #Execute query if table is exists in database
    conn.execute('''Drop table if exists WordFrequnecy;''')

    #Create table in database
    conn.execute('''CREATE TABLE WordFrequnecy
       (word TEXT PRIMARY KEY     NOT NULL,
        Countword INT NOT NULL);''')

#Created a function in order to insert the recorxd into the database table created
def seo_insertRecord(word, cnts):
    conn = sqlite3.connect('E:\\vijay\\seo_program\\seo\\program\\Project\\Project.db')
    conn.execute("INSERT INTO WordFrequnecy VALUES('" + word + "'," + str(cnts) + ")");
    conn.commit()
    conn.close()

#Created a function in order to create a excel sheet and insert the data along with the chart    
def seo_createExcelAndChart():

    #Create workbook
    wb = xlsxwriter.Workbook("E:\\vijay\\seo_program\\seo\\program\\Project\\project.xlsx")
    ws = wb.add_worksheet()
    bold = wb.add_format({'bold': 1})
    
    ws.write("A1","Word", bold)
    ws.write("B1","Total", bold)

    conn = sqlite3.connect('E:\\vijay\\seo_program\\seo\\program\\Project\\Project.db')
    cursor = conn.execute("SELECT * from WordFrequnecy")
    cnt = 2
    for row in cursor:
        ws.write("A"+str(cnt),row[0])
        ws.write("B"+str(cnt),row[1])
        cnt += 1

    #Create chart

    chart1 = wb.add_chart({'type':'line'})
    chart1.add_series({'categories': '=Sheet1!$A$2:$A$5',
                       'data_labels': {'value': True},
                       'values':'=Sheet1!$B2:$B5',
                       'line':   {'color': 'blue'},
                       'marker': {'type': 'square', 'size,': 5, 'border': {'color': 'red'}, 'fill':{'color': 'yellow'}}})

    chart1.set_title ({'name': 'Results of words analysis'})
    chart1.set_x_axis({'name': 'Words'})
    chart1.set_y_axis({'name': 'Occurence of Word'})
    chart1.set_style(10)

    ws.insert_chart("D5", chart1, {'x_offset': 25, 'y_offset': 10})

    wb.close()

#Main functon which consles or completes the rest of the project with the requirements
def main():

    data = read_data("E:\\vijay\\seo_program\\seo\\program\\Project\\PageURL.xlsx")
    siteURL = data["URL1"][0] #geturl from excel file
    keyword1 = data["URL1"][1]
    keyword2 = data["URL1"][2]
    keyword3 = data["URL1"][3]
    keyword4 = data["URL1"][4]
    print(siteURL)

    
    lstWords = {keyword1[0],keyword2[0],keyword3[0],keyword4[0]}

    
    page = seo_fetch_page(siteURL[0])
    
    soup = BeautifulSoup(page)

    for script in soup(["script", "style"]):
        script.extract()    # rip it out
    
    words = soup.get_text().split()
    words.sort()

    #Create table
    seo_createTable()
    
    for w in lstWords:
        print(w, words.count(w))
        seo_insertRecord(w,words.count(w))

    seo_createExcelAndChart()
    
if __name__ == "__main__":
    main()
    
