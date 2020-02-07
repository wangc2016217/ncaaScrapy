from bs4 import BeautifulSoup as bs
import re
from contextlib import suppress
import traceback
from dateutil import parser
import requests
import xlsxwriter
from difflib import SequenceMatcher
from datetime import datetime, date
from datetime import timedelta
from webauto_base import webauto_base
import time
import math
espnArray =  []
collageArray = []
spreadArray = []

def get_espn():
    try:        
        print("Request to https://www.espn.com/mens-college-basketball/standings...")
        url = "https://www.espn.com/mens-college-basketball/standings"
        #Get the Page Content String from the Espn      
        pageString = requests.get(url).text
        soup = bs(pageString, 'html.parser')

        sections = soup.findAll('section', {'class':'ResponsiveTable'})

        for section in sections:
            conference_name = list(section.children)[0].get_text()

            tr = section.findChildren("tr", {'class','Table__TR--sm'})

            for i in range(int(len(tr)/2)):                
                team_name = tr[i].find("abbr")
                point = tr[int(len(tr)/2)+i].findChildren("span", {'class', 'stat-cell'})
                if team_name is not None:                    
                    espnArray.append({
                        "conference_name" : conference_name,
                        "team_name" : team_name['title'],
                        "abbr" : team_name.get_text(),
                        "c_w_l": point[0].get_text(),
                        "c_gb": point[1].get_text(),
                        "c_pct": point[2].get_text(),
                        "o_w_l": point[3].get_text(),
                        "o_pct": point[4].get_text(),
                        "o_home": point[5].get_text(),
                        "o_away": point[6].get_text(),
                        "o_strk": point[7].get_text(),
                    })        
    except Exception as e:
        print(str(e))
        traceback.print_exc()

def get_colleage():
    try:        
        print("Request to https://www.espn.com/mens-college-basketball/bpi/_/view/overview/sort/sospastrank/dir/asc...")
        url = "https://www.espn.com/mens-college-basketball/bpi/_/view/overview/sort/sospastrank/dir/asc"
        pg = 1
        while pg < 100:
            try:
                url = "https://www.espn.com/mens-college-basketball/bpi/_/view/overview/sort/sospastrank/page/" + str(pg) + "/dir/asc"
                page = requests.get(url).text
                soup = bs(page, 'html.parser')

                no_data = soup.find('div', {'class' : 'no-data-available'})
                if no_data is not None and no_data.get_text() == "No data available.":
                    break
                
                trs = soup.findAll('tr')

                for tr in trs:
                    try:
                        td = tr.findChildren('td')
                        if len(td) == 8:
                            collageArray.append({
                                'team_name' : td[1].find('span',{'class','team-names'}).get_text(),
                                'abbr' : td[1].find('abbr').get_text(),
                                'conf' : '',
                                'bpi_rk' : td[4].get_text(),
                                'sos_rk' : td[5].get_text(),
                                'sor_rk' : td[6].get_text(),
                            })
                    except:
                        pass                   

            except Exception as e:
                print(str(e))
                pass
            pg = pg + 1

    except Exception as e:
        print(str(e))
        traceback.print_exc()

class get_spread(webauto_base):
    def __del__(self):
        super().__init__()

    def automate(self):
        try:
            print("Automation to get spreads...")

            url = "https://www.sportsbookreview.com/betting-odds/ncaa-basketball/"
            self.start_browser(False)
            self.navigate(url)

            self.delay_me(3)            

            xpath = "//div[@id='bettingOddsGridContainer']/div[3]/*"
            bettingContainer = self.browser.find_elements_by_xpath(xpath)
                      
            # #print(day)
            xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//div[@data-vertical-sbid='time']/div//span"
            p_time = self.browser.find_elements_by_xpath(xpath)
            #print(len(p_time))

            # #print(len(p_time))

            xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section//a//span"
            teams = self.browser.find_elements_by_xpath(xpath)

            #print(len(teams))
            xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='-1']//span[@data-cy='odd-grid-opener-league']"
            opener = self.browser.find_elements_by_xpath(xpath)

            bookmarker = None
            five_times = None
            bovada = None

            count = 0
            while True:
                xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='93']//span[@data-cy='odd-grid-league']"
                bookmarker = self.browser.find_elements_by_xpath(xpath)
                                
                xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='19']//span[@data-cy='odd-grid-league']"
                five_times = self.browser.find_elements_by_xpath(xpath)

                xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='1618']//span[@data-cy='odd-grid-league']"
                bovada = self.browser.find_elements_by_xpath(xpath)
                                
                if len(opener) == len(bookmarker) or len(opener) == len(five_times) or len(opener) == len(bovada):
                    break

                count = count + 1
                if count > 100 and (len(bookmarker) == 0 or len(five_times) == 0 or len(bovada) == 0):
                    break                               
                
            if len(bookmarker) == 0 or len(five_times) == 0 or len(bovada) == 0:
                xpath = "//i[@class='sbr-icon-chevron-right']"
                right = self.browser.find_element_by_xpath(xpath)
                right = right.find_element_by_xpath("..")
                right.click()

                self.delay_me(3)

                if len(bookmarker) == 0:
                    xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='93']//span[@data-cy='odd-grid-league']"
                    bookmarker = self.browser.find_elements_by_xpath(xpath)
                                
                if len(five_times) == 0:
                    xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='19']//span[@data-cy='odd-grid-league']"
                    five_times = self.browser.find_elements_by_xpath(xpath)

                if len(bovada) == 0:
                    xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='1618']//span[@data-cy='odd-grid-league']"
                    bovada = self.browser.find_elements_by_xpath(xpath)

                right.click()

                self.delay_me(3)

                if len(bookmarker) == 0:
                    xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='93']//span[@data-cy='odd-grid-league']" #93
                    bookmarker = self.browser.find_elements_by_xpath(xpath)
                                
                if len(five_times) == 0:
                    xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='19']//span[@data-cy='odd-grid-league']"
                    five_times = self.browser.find_elements_by_xpath(xpath)

                if len(bovada) == 0:
                    xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section[@data-vertical-sbid='1618']//span[@data-cy='odd-grid-league']"
                    bovada = self.browser.find_elements_by_xpath(xpath) 

            print("Processing...")
            bettingDate = None
            i = 0
            for divs in bettingContainer:
                try:
                    bettingDate = divs.find_element_by_xpath("./div/div/span")
                    if bettingDate.text != "Box Scores":
                        bettingDate = bettingDate.text
                        #print(i)
                        continue

                    if bettingDate.text == "Box Scores":
                        continue
                    
                except:
                    pass
                try:
                    length = len(divs.find_elements_by_xpath("./div"))
                    if length >= 2 and bettingDate is not None:
                        
                        
                        point_A_T = ""
                        point_B_T = ""                        
                        diff_A = 0
                        diff_B = 0
                        if length == 3:
                            points = divs.find_elements_by_xpath("./div[2]/div/div")
                            try:                                
                                point_A_T = int(points[len(points)-1].text.split("\n")[0])
                                point_B_T = int(points[len(points)-1].text.split("\n")[1])
                                diff_A = point_A_T - point_B_T
                                diff_B = point_B_T - point_A_T
                            except:
                                pass                            

                        opener_f = "-"
                        opener_s = "-"
                        try:
                            opener_val = opener[i].find_elements_by_xpath("./span")
                            opener_f = opener_val[0].text
                            opener_s = opener_val[1].text
                        except:
                            pass

                        bookmarker_f = "-"
                        bookmarker_s = "-"
                        try:
                            bookmarker_val = bookmarker[i].find_elements_by_xpath("./span")
                            bookmarker_f = bookmarker_val[0].text
                            bookmarker_s = bookmarker_val[1].text
                        except:
                            pass
                        

                        five_times_f = "-"
                        five_times_s = "-"
                        try:
                            five_times_val = five_times[i].find_elements_by_xpath("./span")
                            five_times_f = five_times_val[0].text
                            five_times_s = five_times_val[1].text
                        except:
                            pass

                        bovada_f = "-"
                        bovada_s = "-"
                        try:
                            bovada_val = bovada[i].find_elements_by_xpath("./span")
                            bovada_f = bovada_val[0].text
                            bovada_s = bovada_val[1].text
                        except:
                            pass

                        spreadArray.append({
                            'team' : teams[i].text,
                            'opener' : opener_f.replace('PK','0').replace('½','.5'),
                            'opener_odds': opener_s.replace('PK','0').replace('½','.5'),
                            'bookmarker' : bookmarker_f.replace('PK','0').replace('½','.5'),
                            'bookmarker_odds' : bookmarker_s.replace('PK','0').replace('½','.5'),
                            'five_times' : five_times_f.replace('PK','0').replace('½','.5'),
                            'five_times_odds' : five_times_s.replace('PK','0').replace('½','.5'),
                            'bovada' : bovada_f.replace('PK','0').replace('½','.5'),
                            'bovada_odds' : bovada_s.replace('PK','0').replace('½','.5'),
                            'date': bettingDate,
                            'point': point_A_T,
                            'diff': diff_A
                        })
                        
                        opener_f = "-"
                        opener_s = "-"
                        try:
                            opener_val = opener[i+1].find_elements_by_xpath("./span")
                            opener_f = opener_val[0].text
                            opener_s = opener_val[1].text
                        except:
                            pass

                        bookmarker_f = "-"
                        bookmarker_s = "-"
                        try:
                            bookmarker_val = bookmarker[i+1].find_elements_by_xpath("./span")
                            bookmarker_f = bookmarker_val[0].text
                            bookmarker_s = bookmarker_val[1].text
                        except:
                            pass
                        

                        five_times_f = "-"
                        five_times_s = "-"
                        try:
                            five_times_val = five_times[i+1].find_elements_by_xpath("./span")
                            five_times_f = five_times_val[0].text
                            five_times_s = five_times_val[1].text
                        except:
                            pass

                        bovada_f = "-"
                        bovada_s = "-"
                        try:
                            bovada_val = bovada[i+1].find_elements_by_xpath("./span")
                            bovada_f = bovada_val[0].text
                            bovada_s = bovada_val[1].text
                        except:
                            pass

                        spreadArray.append({
                            'team' : teams[i+1].text,
                            'opener' : opener_f.replace('PK','0').replace('½','.5'),
                            'opener_odds': opener_s.replace('PK','0').replace('½','.5'),
                            'bookmarker' : bookmarker_f.replace('PK','0').replace('½','.5'),
                            'bookmarker_odds' : bookmarker_s.replace('PK','0').replace('½','.5'),
                            'five_times' : five_times_f.replace('PK','0').replace('½','.5'),
                            'five_times_odds' : five_times_s.replace('PK','0').replace('½','.5'),
                            'bovada' : bovada_f.replace('PK','0').replace('½','.5'),
                            'bovada_odds' : bovada_s.replace('PK','0').replace('½','.5'),
                            'date': bettingDate,
                            'point': point_B_T,
                            'diff': diff_B
                        })
                        i = i + 2
                except:
                    pass

            print("getting point ...")
            yesterday = datetime.strftime(datetime.now() - timedelta(1), '%Y%m%d')
            url = "https://www.sportsbookreview.com/betting-odds/ncaa-basketball/?date=" + yesterday

            xpath = "/html/body/div[1]/div/div/div/section/div/div[2]/div[1]/span[2]"
            self.click_element(xpath, 3, 0)

            self.delay_me(3)            
            
            xpath = "//div[@id='bettingOddsGridContainer']/div[3]/*"
            self.wait_present(xpath, 30)
            bettingContainer = self.browser.find_elements_by_xpath(xpath)                                
            
            xpath = "//div[@id='bettingOddsGridContainer']//div//div//div//section//a//span"
            teams = self.browser.find_elements_by_xpath(xpath)

            bettingDate = None
            i = 0
            for divs in bettingContainer:
                try:
                    bettingDate = divs.find_element_by_xpath("./div/div/span")
                    if bettingDate.text != "Box Scores":
                        bettingDate = bettingDate.text
                        #print(i)
                        continue

                    if bettingDate.text == "Box Scores":
                        continue
                    
                except:
                    pass
                try:
                    length = len(divs.find_elements_by_xpath("./div"))
                    if length >= 2 and bettingDate is not None:
                                                
                        point_A_T = ""
                        point_B_T = ""                        
                        diff_A = 0
                        diff_B = 0

                        teamNameA = ""
                        teamNameB = ""
                        if length == 3:
                            points = divs.find_elements_by_xpath("./div[2]/div/div")
                            teamNameA = divs.find_element_by_xpath("./div[1]/section/div[1]/div[1]/div/a/span").text
                            teamNameB = divs.find_element_by_xpath("./div[1]/section/div[1]/div[2]/div/a/span").text
                            try:                                
                                point_A_T = int(points[len(points)-1].text.split("\n")[0])
                                point_B_T = int(points[len(points)-1].text.split("\n")[1])
                                diff_A = point_A_T - point_B_T
                                diff_B = point_B_T - point_A_T
                            except:
                                pass

                        if spreadArray[i]["point"] == "" and teamNameA == spreadArray[i]["team"]:
                            spreadArray[i]["point"] = point_A_T
                            spreadArray[i]["diff"] = diff_A

                        if spreadArray[i+1]["point"] == "" and teamNameB == spreadArray[i+1]["team"]:
                            spreadArray[i+1]["point"] = point_B_T
                            spreadArray[i+1]["diff"] = diff_B
                        i = i + 2
                except:
                    pass
            self.quit_browser()
        except Exception as e:
            print(str(e))
            traceback.print_exc()

def findDay():
    return date.today().weekday()        

def get_conf(team_name):

    try:
        tmp = espnArray[0]
        pt = 0
        for espn in espnArray:
            if pt < SequenceMatcher(None, espn['team_name'].lower(), team_name.lower()).ratio():
                pt = SequenceMatcher(None, espn['team_name'].lower(), team_name.lower()).ratio()
                tmp = espn

        return tmp
    except:
        return None

def get_coll(team_name):
    try:
        tmp = collageArray[0]
        pt = 0
        for collage in collageArray:
            if pt < SequenceMatcher(None, collage['team_name'].lower(), team_name.lower()).ratio():
                pt = SequenceMatcher(None, collage['team_name'].lower(), team_name.lower()).ratio()
                tmp = collage

        return tmp
    except:
        return None

def getCurrentDate():
    return datetime.today().day    

def getCurrentMonth():
    return datetime.today().strftime("%b")

def make_excel():
    print("Creating...")
    today = datetime.now()

    dayNumber = findDay()
    filename = "result_" + today.strftime("%Y%m%d%H%M%S") + ".xlsx"
    #filename = "result.xlsx"
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook._get_sheet_index('NCAA Teams')    
    #worksheet = workbook.get_worksheet_by_name('NCAA Teams')
    

    #worksheet = workbook.add_worksheet('NCAA Teams')
    
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 0, "Team", bold)
    worksheet.write(0, 1, "Conference", bold)
    worksheet.write(0, 2, "W-L (Conference)", bold)
    worksheet.write(0, 3, "GB", bold)
    worksheet.write(0, 4, "PCT", bold)
    worksheet.write(0, 5, "W-L (Overall)", bold)
    worksheet.write(0, 6, "PCT", bold)
    worksheet.write(0, 7, "Home", bold)
    worksheet.write(0, 8, "PCT", bold)
    worksheet.write(0, 9, "Away", bold)
    worksheet.write(0, 10, "PCT", bold)
    worksheet.write(0, 11, "STRK", bold)
                        
    index = 1
    for espn in espnArray:
        worksheet.write(index, 0, espn['team_name'])
        worksheet.write(index, 1, espn['conference_name'])
        worksheet.write(index, 2, espn['c_w_l'])
        worksheet.write(index, 3, espn['c_gb'])
        worksheet.write(index, 4, espn['c_pct'])
        worksheet.write(index, 5, espn['o_w_l'])
        worksheet.write(index, 6, espn['o_pct'])
        worksheet.write(index, 7, espn['o_home'])
        worksheet.write(index, 9, espn['o_away'])
        worksheet.write(index, 11, espn['o_strk'])

        result = 0
        try:
            array = espn["o_home"].split('-')                   
            result = round(int(array[0])/(int(array[0])+int(array[1])) * 100)
        except:
            result = 0
            pass
        worksheet.write(index, 8, str(result) + "%")

        result = 0
        try:
            array = espn["o_away"].split('-')                   
            result = round(int(array[0])/(int(array[0])+int(array[1])) * 100)
        except:
            result = 0
            pass
        worksheet.write(index, 10, str(result) + "%")

        index = index + 1

    worksheet = workbook.add_worksheet('Spread')
    
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 0, "Date", bold)
    worksheet.write(0, 1, "Team", bold)
    worksheet.write(0, 2, "Conf", bold)
    worksheet.write(0, 3, "Spread (Opener)", bold)
    worksheet.write(0, 4, "Odds", bold)
    worksheet.write(0, 5, "Spread (BookMaker)", bold)
    worksheet.write(0, 6, "Odds", bold)
    worksheet.write(0, 7, "Spread (5 Dimes)", bold)
    worksheet.write(0, 8, "Odds", bold)
    worksheet.write(0, 9, "Spread (Bovada)", bold)
    worksheet.write(0, 10, "Odds", bold)
    worksheet.write(0, 11, "Away/Home Overall Record", bold)
    worksheet.write(0, 12, "Percentage", bold)
    worksheet.write(0, 13, "W-L (Overall) PCT", bold)
    worksheet.write(0, 14, "STRK", bold)
    worksheet.write(0, 15, "BPI Rank", bold)
    worksheet.write(0, 16, "SOS Rank", bold)
    worksheet.write(0, 17, "SOR Rank", bold)
    worksheet.write(0, 18, "Score", bold)
    worksheet.write(0, 19, "P.D", bold)
    worksheet.write(0, 20, "Away 30% below", bold)
    worksheet.write(0, 21, "Home 70% above", bold)
    worksheet.write(0, 22, "Sharp-Square", bold)
    worksheet.write(0, 23, "Wager", bold)

    espn['o_strk']
    
    green_format1 = workbook.add_format()
    green_format1.set_pattern(1)  # This is optional when using a solid fill.
    green_format1.set_bg_color('green')

    tomato_format2 = workbook.add_format()
    tomato_format2.set_pattern(1)  # This is optional when using a solid fill.
    tomato_format2.set_bg_color('#ff6347')

    magenta_format = workbook.add_format()
    magenta_format.set_pattern(1)  # This is optional when using a solid fill.
    magenta_format.set_bg_color('magenta')   
    index = 1
    flag = False
    for spread in spreadArray: 
               
        worksheet.write(index, 0, spread['date'])
        worksheet.write(index, 1, spread['team'])
        first_A = get_conf(spread['team'])
        first_B = get_coll(spread['team'])

        if first_A is not None:
            worksheet.write(index, 2, first_A["conference_name"])
            worksheet.write(index, 13, first_A["o_pct"])
            worksheet.write(index, 14, first_A["o_strk"])

            if index % 3 == 1:
                worksheet.write(index, 11, first_A["o_away"])
                result = 0                
                try:
                    array = first_A["o_away"].split('-')                   
                    result = round(int(array[0])/(int(array[0])+int(array[1])) * 100)
                except:
                    result = 0
                    pass

                # if result >= 70:
                #     worksheet.write(index, 12, str(result) + "%",green_format1)
                if result <= 30:
                    worksheet.write(index, 12, str(result) + "%",tomato_format2)
                    flag = True
                else:
                    worksheet.write(index, 12, str(result) + "%")
                    flag = False
            else:
                worksheet.write(index, 11, first_A['o_home'])

                result = 0
                try:
                    array = first_A["o_home"].split('-')
                    result = round(int(array[0])/(int(array[0])+int(array[1])) * 100)                    
                except:
                    result = 0
                    pass
                if result >= 70:
                    worksheet.write(index, 12, str(result) + "%",green_format1)
                    worksheet.write(index, 21, 1)
                # elif result <= 30:
                #     worksheet.write(index, 12, str(result) + "%",tomato_format2)
                else:
                    worksheet.write(index, 12, str(result) + "%")
                    worksheet.write(index, 21, 0)
                try:
                    worksheet.write(index, 22, float(spread['bovada']) - float(spread['bookmarker']))
                except:
                    worksheet.write(index, 22, 0)

                worksheet.write(index, 23, "no")

                if flag == True:
                    worksheet.write(index, 20, 1)
                    try:
                        if result >= 70 and float(spread['bovada']) - float(spread['bookmarker']) > 0:
                            worksheet.write(index, 23, "yes")
                    except:
                        pass
                else:
                    worksheet.write(index, 20, 0)

        if first_B is not None:
            worksheet.write(index, 15, first_B['bpi_rk'])
            worksheet.write(index, 16, first_B['sos_rk'])
            worksheet.write(index, 17, first_B['sor_rk'])
        worksheet.write(index, 18, spread['point'])

        if spread['diff'] > 0:
            worksheet.write(index, 19, spread['diff'])
        
            
        worksheet.write(index, 3, spread['opener'])
        worksheet.write(index, 4, spread['opener_odds'])

        if spread['bookmarker'] == spread['bovada']:
            worksheet.write(index, 5, spread['bookmarker'])
        else:
            worksheet.write(index, 5, spread['bookmarker'],magenta_format)
            
        worksheet.write(index, 6, spread['bookmarker_odds'])
        worksheet.write(index, 7, spread['five_times'])
        worksheet.write(index, 8, spread['five_times_odds'])

        if spread['bookmarker'] == spread['bovada']:
            worksheet.write(index, 9, spread['bovada'])
        else:
            worksheet.write(index, 9, spread['bovada'],magenta_format)
        worksheet.write(index, 10, spread['bovada_odds'])
        
        index = index + 1
        if index % 3 == 0:
            index = index + 1      
    workbook.close()
    pass

def main():
    get_espn()
    get_colleage()
    spread = get_spread()
    spread.automate()
    make_excel()    
     
main()
