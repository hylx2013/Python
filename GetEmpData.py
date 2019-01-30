# coding: utf-8
import requests
import xlsxwriter
from bs4 import BeautifulSoup


headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.8",
    "Connection": "keep-alive",
    "Cookie": "_cod=9.10.11.12.13.14.15; csd=18; UM_distinctid=1657a4e4eba212-09ff84a48b2a91-9393265-100200-1657a4e4ebb6f3; NAMEID=1546048659; CNZZDATA1140722=cnzz_eid%3D1379050525-1535354373-http%253A%252F%252Fwww.hnzyfy.com%252F%26ntime%3D1547691799",
    "Referer": "http://www.hnzyfy.com/kssz/class/?170.html",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.98 Safari/537.36 LBBROWSER"
}

url = 'http://www.hnzyfy.com/kssz/class/?'
personurl="http://www.hnzyfy.com/kssz/html/?"

deplist=[18,22,179,13,176,171,175,170,196,162,173,169,177,160,159,168,158,49,167,174,166,141,50,172,48,43,41,35,52,47,40,34,26,21,53,46,39,33,25,20,54,38,36,32,19,55,44,31,29,17,56,28,22,18]
#deplist=[18]
emplist=[]
empObjList=[]

def getEmplist(urlparam):
    response = requests.get(urlparam, headers=headers)
    response.encoding = ('utf8')
    #print(response.text)
    soup = BeautifulSoup(response.text, 'lxml')
   # print(len(alist))
    for td in soup.find_all("td",class_="cpquery"):
       f_empurl=""
       f_imgurl=""
       if not td.contents is None:
           if not td.find("a") is None:
               f_empurl = td.find('a').get("href")
           if not td.find("img") is None:
               f_imgurl = td.find('img').get("src").replace("../..","http://www.hnzyfy.com")
               #print(f_imgurl)
           emplist.append({"empurl": f_empurl, "imgurl": f_imgurl})
    if not soup.find("td",class_="pages") is None:
        #print('获取第二页'+str(urlparam))
        urlparam2=urlparam+'&page=2&showtj=&showhot=&key=&ksid='
        response2 = requests.get(urlparam2, headers=headers)
        response2.encoding = ('utf8')
       # print(response2.text)
        soup2 = BeautifulSoup(response.text, 'lxml')
        for td2 in soup2.find_all("td", class_="cpquery"):
            f_empurl2=""
            f_imgurl2=""
            if not td2.contents is None:
                if not td2.find("a") is None:
                    f_empurl2=td2.find('a').get("href")
                if not td2.find("img") is None:
                    f_imgurl2=td2.find('img').get("src").replace("../..","http://www.hnzyfy.com")
                emplist.append({"empurl":f_empurl2,"imgurl":f_imgurl2})
def generate_excel(rec_data):
    workbook = xlsxwriter.Workbook('C://Users//Hey//Desktop//emp.xlsx')
    worksheet = workbook.add_worksheet()
    bold_format = workbook.add_format({'bold': True})
    worksheet.write('A1', 'name', bold_format)
    worksheet.write('B1', 'introduction', bold_format)
    worksheet.write('C1', 'title', bold_format)
    worksheet.write('D1', 'imgUrl', bold_format)
    row = 1
    col = 0
    for item in rec_data:
        worksheet.write_string(row, col, item['name'])
        worksheet.write_string(row, col + 1, item['introduction'])
        worksheet.write_string(row, col + 2, str(item['title']))
        worksheet.write_string(row, col + 3, item['imgUrl'])
        row += 1
    workbook.close()

for deptId in deplist:
    urlDept=url+str(deptId)+".html"
    getEmplist(urlDept)

for empUrl in emplist:
    empUrlReplace="http://www.hnzyfy.com/kssz"+empUrl["empurl"].replace("..","")
    response = requests.get(empUrlReplace, headers=headers)
    response.encoding = ('utf8')
    soup=BeautifulSoup(response.text, 'lxml')
    tds=soup.find_all("td",class_="cpprop")
    if not tds is None:
        introduction=soup.find("td",class_="cpintro")
        introductiontext="";
        if not introduction is None:
            #print(introduction)
            if not introduction.find("p") is None:
                introductiontext=introduction.find("p").get_text()
        print(tds)
        print(empUrl)
        if len(tds)>0 and  not empUrl is None:
            empObj={"name":tds[0].get_text(),"introduction":introductiontext,"title":tds[2].get_text(),"imgUrl":empUrl["imgurl"]}
        #empObj=Employee(tds[0].get_text(),introductiontext,tds[2].get_text(),imgUrl)
        empObjList.append(empObj)

generate_excel(empObjList)

