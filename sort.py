#!/usr/bin/env python3

import random
import re
import collections
import requests
from html.parser import HTMLParser
import openpyxl
import datetime
import time
import locale
import os
import settings

#192.168.0.248', '-', 'servers', '[25/Mar/2019:01:18:25', '+0500]', '"CONNECT', 'https://195.122.177.135:443/', 'HTTP/1.0"', '200', '856
WEBTIMEOUT=60 # время в минутах счётчика посещений
COUNT,SIZEB,UXTIME,VISITS=(0,1,2,3)
IP,USER,DATE,LINK,BYTES = (0,2,3,6,9)

def main():
    locale.setlocale(locale.LC_ALL, '')
    requests.packages.urllib3.disable_warnings()

    filename=fileOpen()
    f=open(filename, 'r')
    rows = f.read().splitlines()

    row=0 # счётчик строк лога
    lst={} # словарь пользователя
    users=set() # список пользователей
    dates=set() # список дат по дням
    mindate=999999999999999
    maxdate=0

    SplitStrings(rows,users,dates)
    try:
        username=input("Enter Username:")
    except KeyboardInterrupt:
        print('Break...')
    except Exception:
        print('ok....')

    print(username)   
    if username not in users:
        print("Incorrect username")
        exit()

    while row<len(rows):
        try:
            user = rows[row][USER]
            link = rows[row][LINK]
            link = http_parse(link)
            byte = int(rows[row][BYTES])
            uxtime = TimeToTs(rows[row][DATE])

            mindate=uxtime if uxtime<mindate else mindate
            maxdate=uxtime if uxtime>maxdate else maxdate
            row+=1

            if username!="All":
                if user!=username:
                    continue

            dbres=lst.get(user) 
            if dbres is None:
                #First user record - Count,Size,UnixTime,Visits
                links={}
                links[link]=[1,byte,uxtime,1]
                lst[user]=links
            else:
                dbres = lst.pop(user)
                lst[user] = CheckLinkExists(dbres,link,byte,uxtime)
        except EOFError:
            break

    wb=openpyxl.Workbook()

    for user in sorted(lst.keys(),reverse=True):
        if user in settings.EXCLUDE:
            continue
        data = lst[user]
        print("___ User: ",user," Items:,",len(data),"______________________________")
        sheet = xlsHead(wb,user,mindate,maxdate)
        lenlink,lensize=0,0
        xrow=3 # Стартовая строка для записи логов

        for lnk,opts in sorted(data.items(), reverse=True, key=lambda x: x[1][1]):
            print("Link: {0}  Options: {1}  Traff:{2}".format(lnk,opts,traf(opts[SIZEB])))
            if opts[SIZEB] == 0:
                continue
            if len(lnk) > lenlink:
                lenlink = len(lnk)
            if len(traf(opts[SIZEB])) > lensize:
                lensize = len(traf(opts[SIZEB]))
            xlsInsert(sheet,xrow,2,lnk,traf(opts[SIZEB]),opts[COUNT],opts[VISITS])
            xrow+=1
        xlsSetColumn(sheet,lenlink,lensize,10,10)
    print("MinDate",datetime.datetime.fromtimestamp(mindate)," MaxDate: ",datetime.datetime.fromtimestamp(maxdate))
    wb.save('example.xlsx')
    return


def fileOpen():
    xfiles={}
    ifile=0
    files = [x for x in os.listdir(".") if x.endswith(".log")]
    if not files:
        filename=input("Enter Log Filename: ")
        print(filename,len(filename))
        if len(filename)==0:
            exit()
        return(filename)
    for fl in sorted(files,reverse=True):
        xfiles[ifile]=fl
        print("{0}: {1}".format(ifile,fl))
        ifile+=1
    try:
        index=int(input("Please enter Index of Filename:"))
        filename=xfiles[index]
    except KeyError:
        print("Enter correct Index of Filename")
        exit()
    print("Use ",filename)
    return(filename)


def xlsInsert(sheet,xrow,xcol,link,traf,req,visits):
    """
    sheet - xls sheet,(xrow,xcol) - row + column, link - link of resource, traf - traffic from link, req - numbers os requests from browser, vivists - visits in hour
    """
    cell = sheet.cell(row = xrow, column = xcol)
    cell.value=link
    xcol+=1
    cell = sheet.cell(row = xrow, column = xcol)
    cell.value=traf
    xcol+=1
    cell = sheet.cell(row = xrow, column = xcol)
    cell.value=req
    xcol+=1
    cell = sheet.cell(row = xrow, column = xcol)
    cell.value=visits
    #if lnk.startswith("http:"):
    #    print(GetTitle(lnk))
    return


def xlsHead(wb,listname,mindate,maxdate):
    wb.create_sheet(listname, index = 0)
    sheet = wb[listname]
    xrow=3
    sheet.merge_cells("B1:E1")
    #font = openpyxl.styles.Font(name='Arial', size=24, italic=True, color='FF0000')
    font = openpyxl.styles.Font(bold=True)
    sheet['B2'].font = font
    sheet['C2'].font = font
    sheet['D2'].font = font
    sheet['E2'].font = font
    cell = sheet.cell(row = 1, column = 2)
    cell.value="Date: "+str(datetime.datetime.fromtimestamp(mindate))+" - "+str(datetime.datetime.fromtimestamp(maxdate))
    cell = sheet.cell(row = 2, column = 2)
    cell.value="Link"
    cell = sheet.cell(row = 2, column = 3)
    cell.value="Size"
    cell = sheet.cell(row = 2, column = 4)
    cell.value="Requests"
    cell = sheet.cell(row = 2, column = 5)
    cell.value="Req/h"
    return sheet


def xlsSetColumn(sheet,first,second,third,four):
    sheet.column_dimensions['B'].width=first
    sheet.column_dimensions['C'].width=second
    sheet.column_dimensions['D'].width=third
    sheet.column_dimensions['E'].width=four
    return


def TimeToTs(dtime):
    dtime=dtime[1::]
    return(int(datetime.datetime.strptime(dtime, '%d/%b/%Y:%H:%M:%S').strftime("%s")))


def traf(byte=0):
    if byte>1024*1024:
        byte=byte/1024/1024
        return(str(round(byte,1))+' Mbytes')
    if byte>1024:
        byte=byte/1024
        return(str(round(byte,1))+' Kbytes')
    return(str(byte)+' Bytes')


def CheckLinkExists(dbres,link,byte,uxtime):
    for dblink,options in dbres.items():
        if dblink == link:
            #Count links
            count=options[COUNT]
            count+=1
            #Size of link
            sizeb=options[SIZEB]
            sizeb+=byte
            #Visits
            visits=int(options[VISITS])
            #UxTime
            dbuxtime=int(options[UXTIME])
            if (uxtime-dbuxtime>60*WEBTIMEOUT):
                dbuxtime=uxtime
                visits+=1
            dbres[link]=[count,sizeb,dbuxtime,visits]
            return(dbres)
    else:
        dbres[link]=[1,byte,uxtime,1]
    return(dbres)


def SplitStrings(rows,users,dates):
    print("Len rows:",len(rows))
    for i in range(len(rows)):
        rows[i]=rows[i].split(" ")
        users.add(rows[i][USER])
        date=rows[i][DATE].split(":")[0]
        date=date[1::]
        dates.add(date)
    users.add("All")
    users=list(users)
    users=sorted(users,key=lambda x: x[0])
    dates=list(dates)
    dates=sorted(dates)
    print("Users:",users)
    print("Dates:",dates)
    return


def http_parse(link):
    if link.startswith("http:"):
        link=link.split("//")[1]
        link=link.split("/")[0]
        #link="http://"+link
    if link.startswith("https:"):
        link=link.split("//")[1]
        link=link.split("/")[0]
    if link.endswith(":443"):
        link=link.split(":443")[0]
    return(link)


class MyHTMLParser(HTMLParser):
    def handle_endtag(self, tag):
        if tag == 'title':
            raise StopIteration()
    def handle_data(self, data):
        self.title = data


def GetTitle(url):
    http_proxy  = "http://192.168.1.1:8080"
    https_proxy = "https://192.168.1.1:8080"
    ftp_proxy   = "ftp://192.168.1.1:8080"
    proxyDict = {
                  "http"  : http_proxy,
                  "https" : https_proxy,
                  "ftp"   : ftp_proxy
                }
    try:
        r = requests.get(url,stream=True,proxies=proxyDict,verify=False,timeout=1) # включаем потоковый режим
        data = next(r.iter_content(2048)) # запрашиваем ровно 512 байт, для чтения тега head этого должно хватать, или можно еще увеличить
        #print("!!!!",data,"!!!!")
        parser = MyHTMLParser()
        if r.encoding is None:
            r.close()
            return None
        sdata = data.decode(r.encoding)
        r.close()
    except requests.exceptions.SSLError:
        return None
    except StopIteration:
        return None
    except Exception:
        return None

    try:
        parser.feed(sdata)
    except StopIteration:
        print(parser.title)
    except Exception:
        return None
    return


main()
