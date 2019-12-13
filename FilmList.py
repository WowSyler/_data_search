import requests 
from bs4 import BeautifulSoup
from openpyxl import Workbook
import matplotlib.pyplot as plt
from collections import Counter
import numpy as np
import time


class FilmListXml():
    def SaveList(self):
        try:
            f = open("films.xlsx")
            f.close()
        except IOError:
            site1 = "https://www.wildaboutmovies.com/201"
            site2 = "_movies/"
            dates = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

            data = []
            dataa = []
            

            films = Workbook()
            sheet = films.active

            for allpage in range(6, 9, 1):
                url = site1 + str(allpage) + site2
                print("\nPage:" + url)
                url_oku = requests.get(url)
                soup = BeautifulSoup(url_oku.content, 'html.parser')

                gelenveri = soup.find('article', {'class':'post-grid'})
                dateNow = ""
                filmName = ""
                writeName = ""
                for f in gelenveri.findAll():
                    for date in dates:
                        if f.text.startswith(date):
                            dateNow = f.text
                            print(f.text)
                            
                    alink = f.find('p')
                    #print(alink)
                    if dateNow != "":
                        if f.text.isspace():
                            pass
                        else:
                            filmName = f.text
                        if filmName != None or filmName != "" and filmName != writeName:
                            for date2 in dates:
                                sonuc = f.text.startswith(date2)
                                if sonuc:
                                    break
                            if f.text != None or f.text != "" or f.text != "\n" or f.text != "\n\n":
                                if alink != None:
                                    #print(alink.text)
                                    sheet.append((alink.text,dateNow))
                                    writeName = filmName
                    
            films.save("films.xlsx")
            films.close()
        #finally:
            #f.close()

       




