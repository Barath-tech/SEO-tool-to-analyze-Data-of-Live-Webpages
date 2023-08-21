import sqlite3
from matplotlib import  pyplot as plt
import bs4
import requests
import xlsxwriter
import xlrd
import json

class Webscrap:
    final_data = []
    read_data=[]
    def data_extract(self,soup):
        data = json.loads(str(soup))
        data1 = data["personList"]["personsLists"]
        # print(data1)

        print("********************************************")
        print("The Number of data present in the pages is",len(data1))
        for dt in data1:
            self.first_data = [
                dt.get("rank", ''),
                dt.get("personName", '').strip('&amp'),
                dt.get("finalWorth", ''),
                dt.get("age", ''),
                dt.get("state", ''),
                dt.get("source", ''),
                dt.get("philanthropyScore", '')
            ]
            # print(self.first_data)
            self.final_data.append(self.first_data)
        #print(self.final_data)

    def excel_data(self):
        workbook=xlsxwriter.Workbook("The_Richest_People_in_America.xlsx")
        worksheet=workbook.add_worksheet()
        bold=workbook.add_format({'bold':True})
        worksheet.write('A1','RANK',bold)
        worksheet.write('B1', 'PERSON_NAME', bold)
        worksheet.write('C1', 'FINAL_WORTH', bold)
        worksheet.write('D1', 'AGE', bold)
        worksheet.write('E1', 'STATE', bold)
        worksheet.write('F1', 'SOURCE', bold)
        worksheet.write('G1', 'PHILANTHROPY_SCORE', bold)
        print("************************************************")
        print("EXCEL SHEET CREATED SUCCESSFULLY.....")
        row=1
        col=0
        for data1 in self.final_data:
            worksheet.write(row,col,data1[0])
            worksheet.write(row,col+1,data1[1])
            worksheet.write(row,col+2,data1[2])
            worksheet.write(row,col+3,data1[3])
            worksheet.write(row,col+4,data1[4])
            worksheet.write(row,col+5,data1[5])
            worksheet.write(row,col+6,data1[6])

            row+=1
        print("**********************************************")
        print("DATA WRITTEN ON EXCEL SUCCESSFULLY.....")


        chart1=workbook.add_chart({'type':'column'})
        chart1.add_series({'categories':'=Sheet1!$B$2:$B$50','values':'=Sheet1!$C$2:$C$50'})
        chart1.add_series({'categories':'=Sheet1!$B$2:$B$50','values':'=Sheet1!$G$2:$G$50'})
        chart1.set_title({'name':'RICHEST PEOPLE IN AMERICA'})
        worksheet.insert_chart('K4',chart1)
        workbook.close()
        print("*************************************")
        print("GRAPH OF THE DATA SUCCESSFULLY DRAWN ON EXCEL SHEET")

    def read_excel(self):
        wb=xlrd.open_workbook("The_Richest_People_in_America.xlsx")
        worksheet=wb.sheet_by_name("Sheet1")
        num_rows=worksheet.nrows
        num_cols=worksheet.ncols
        row_review = []
        for current_row in range(0,num_rows,1):
            row_review=[]

            for current_col in range(0,num_cols,1):
                review=worksheet.cell_value(current_row,current_col)

                row_review.append(review)

            self.read_data.append(row_review)
        #print(self.read_data)

        print("*********************** **************")
        print("DATA SUCCESSFULLY READ FROM EXCEL SHEET")

    def data_base(self):

        data_base_values=self.read_data
        con=sqlite3.connect("RPA.db")
        print("*************************************")
        print("DATABASE CONNECTED SUCCESSFULLY..")
        cur = con.cursor()
        listOfTables = cur.execute(
            """SELECT 'RICHEST_PEOPLE_IN_AMERICA' FROM sqlite_master WHERE type='table' ; """).fetchall()

        if listOfTables == []:
            print('Table not found!')
            cur.execute('''CREATE TABLE RICHEST_PEOPLE_IN_AMERICA (
            	Rank INTEGER NOT NULL,
            	Person_name	TEXT,
            	Final_worth INTEGER NOT NULL,
            	Age	INTEGER NOT NULL,
            	State TEXT,
            	Source_value TEXT,
            	Philanthropy_score INTEGER);''')
        else:
            print('Table found!')

        #print(data_base_values)

        cur.executemany("INSERT INTO RICHEST_PEOPLE_IN_AMERICA(Rank,Person_name,Final_worth,Age,State,Source_value,Philanthropy_score) VALUES(?,?,?,?,?,?,?)",data_base_values)
        con.commit()
        print("*************************************")
        print("DATA STORED IN DATABASE SUCCESSFULLY!!")

    def graph(self):
        first_pt=[dt[1] for dt in self.final_data]
        second_pt=[dt[2] for dt in self.final_data]
        #print(first_pt)
        #print(second_pt)
        plt.bar(first_pt,second_pt,color='b')
        plt.legend(["NAMES","FINAL_WORTH"])
        plt.show()
        print("*************************************")
        print("BAR GRAPH DRAWN SUCCESSFULLY..!!")

        plt.scatter(first_pt,second_pt,label='cases', color='b')
        plt.show()
        print("*************************************")
        print("SCATTER PLOT DRAWN SUCCESSFULLY..!!")





try:
    urllink="https://www.forbes.com/forbesapi/person/forbes-400/2021/position/true.json?limit=1000&fields=personName,rank,age,gender,finalWorth,industries,philanthropyScore,state,source,bios,squareImage,uri,status"
    header={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36'
    }
    response=requests.get(url=urllink,headers=header)
    soup=bs4.BeautifulSoup(response.content,"html.parser")
    #print(soup.prettify())
    w=Webscrap()
    w.data_extract(soup)
    print("*************************************")
    print("WEB PAGE EXTRACTED SUCCESSFULLLY!!!")
    w.excel_data()
    w.read_excel()
    w.data_base()
    w.graph()
    print("********************************************************************************")
    print("WEB SCRAPING OF THE 'THE RICHEST PEOPLE OF AMERICA' IS DONE SUCCESSFULLY..!!")
    print("********************************************************************************")

except Exception as e:
    print("Exception occurs as",e,"Please check and run it again...")



