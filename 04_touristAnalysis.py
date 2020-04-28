import requests, datetime, time, os
from bs4 import BeautifulSoup as bs
from openpyxl import Workbook
session_requests = requests.session()
loginUrl = 'https://springtour.app/Ajax/loginadmin'
# payload = {
#     'account': 'xxxx',
#     'password': 'xxxx'
# }
payload = {
    'account': 'xxxx',
    'password': 'xxxx'
}
headers = {
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Origin': 'https://springtour.app',
    'Referer': 'https://springtour.app/Main/loginAdmin',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest'
}
totalTetantCount = 0
date = str(datetime.date.today()).replace('-', '')
folderName = '{} 雲林自由行旅客分析'.format(date)
if not os.path.exists(folderName):
    os.mkdir(folderName)
r = session_requests.post(loginUrl, data=payload, headers=headers)
if r.status_code == requests.codes.ok:
    print('Login Success!')
    hotelCount = 1
    for i in range(1, 33):
        targetUrl = 'https://springtour.app/Gov/statis/{}'.format(i)
        r = session_requests.get(targetUrl, headers = dict(referer = targetUrl))
        soup = bs(r.text, 'html.parser')
        hotelInfos = soup.find_all('td', {'class': 'cal6'})
        for hotelInfo in hotelInfos:
            hotelLink = hotelInfo.find('a').get('href')
            hotelUrl = 'https://springtour.app{}'.format(hotelLink)
            r = session_requests.get(hotelUrl, headers = dict(referer = targetUrl))
            soup = bs(r.text, 'html.parser')
            hotelNameTag = soup.select('#user_panel p')[0]
            hotelName = hotelNameTag.contents[0]
            hotelTenantTotal = 0
            genderDict = {}
            birthPlaceDict = {}
            birthEraDict = {}
            print(hotelName)
            wb = Workbook()
            ws = wb.active
            orderNumbers = soup.find_all('td', {'class': 'cal1'})
            orderInfos = soup.find_all('td', {'class': 'cal6'})
            orderNumbersLst = []
            orderInfosLst = []
            rowCount = 0
            getDataOfOrder = soup.find_all('td')
            # print(len(getDataOfOrder))
            if len(getDataOfOrder) > 0:
                for orderNumber in orderNumbers:
                    orderNumber = orderNumber.text
                    orderNumbersLst.append(orderNumber)
                for orderInfo in orderInfos:
                    orderLink = orderInfo.find('a').get('href')
                    orderUrl = 'https://springtour.app{}'.format(orderLink)
                    orderInfosLst.append(orderUrl)
                orderDict = dict(zip(orderNumbersLst, orderInfosLst))
                for orderKey, orderDetailUrl in orderDict.items():
                    # print(orderKey, len(orderKey))
    # convert start===========================
                    orderDetail = orderDetailUrl
                    r = session_requests.get(orderDetail)
                    soup = bs(r.text, 'html.parser')
                    # get id count =============================================
                    td_tags = soup.find_all('td', {'class': 'cal6'})
                    idCount = 1
                    for td in td_tags:
                        for data in td:
                            # print(data)
                            if str(data).startswith('<div>') and idCount % 2 == 0:
                                data = str(data).replace('<div>', '').replace('</div>', '').replace('<p>', '').replace('</p>', '').replace('\n', '')
                                # print(idCount, data)
                                if data[0] not in birthPlaceDict:
                                    birthPlaceDict[data[0]] = 0
                                    birthPlaceDict[data[0]]  += 1
                                else:
                                    birthPlaceDict[data[0]]  += 1
                                if data[1] not in genderDict:
                                    genderDict[data[1]] = 0
                                    genderDict[data[1]] += 1
                                else:
                                    genderDict[data[1]] += 1
                                idCount += 1
                            else:
                                idCount += 1
                    idCount = 1
                    # id count end ==============================================
                    birthEra_tags = soup.find_all('td', {'class': 'cal7'})
                    for birthEra in birthEra_tags:
                        for data in birthEra:
                            if str(data).startswith('<div>'):
                                data = str(data).replace('<div>', '').replace('</div>', '').replace('<p>', '').replace('</p>', '').replace('\n', '')
                                birthYear = int(str(data.split('/',1)[0]))
                                if (108 - birthYear) <= 10:
                                    if '10歲(含)以下' not in birthEraDict:
                                        birthEraDict['10歲(含)以下'] = 0
                                        birthEraDict['10歲(含)以下'] += 1
                                    else:
                                        birthEraDict['10歲(含)以下'] += 1
                                elif (108 - birthYear) > 10 and (108 - birthYear) <= 20:
                                    if '11-20歲' not in birthEraDict:
                                        birthEraDict['11-20歲'] = 0
                                        birthEraDict['11-20歲'] += 1
                                    else:
                                        birthEraDict['11-20歲'] += 1
                                elif (108 - birthYear) > 20 and (108 - birthYear) <= 30:
                                    if '21-30歲' not in birthEraDict:
                                        birthEraDict['21-30歲'] = 0
                                        birthEraDict['21-30歲'] += 1
                                    else:
                                        birthEraDict['21-30歲'] += 1
                                elif (108 - birthYear) > 30 and (108 - birthYear) <= 40:
                                    if '31-40歲' not in birthEraDict:
                                        birthEraDict['31-40歲'] = 0
                                        birthEraDict['31-40歲'] += 1
                                    else:
                                        birthEraDict['31-40歲'] += 1
                                elif (108 - birthYear) > 40 and (108 - birthYear) <= 50:
                                    if '41-50歲' not in birthEraDict:
                                        birthEraDict['41-50歲'] = 0
                                        birthEraDict['41-50歲'] += 1
                                    else:
                                        birthEraDict['41-50歲'] += 1
                                elif (108 - birthYear) > 50 and (108 - birthYear) <= 60:
                                    if '51-60歲' not in birthEraDict:
                                        birthEraDict['51-60歲'] = 0
                                        birthEraDict['51-60歲'] += 1
                                    else:
                                        birthEraDict['51-60歲'] += 1
                                elif (108 - birthYear) > 60 and (108 - birthYear) <= 70:
                                    if '61-70歲' not in birthEraDict:
                                        birthEraDict['61-70歲'] = 0
                                        birthEraDict['61-70歲'] += 1
                                    else:
                                        birthEraDict['61-70歲'] += 1
                                elif (108 - birthYear) > 70 and (108 - birthYear) <= 80:
                                    if '71-80歲' not in birthEraDict:
                                        birthEraDict['71-80歲'] = 0
                                        birthEraDict['71-80歲'] += 1
                                    else:
                                        birthEraDict['71-80歲'] += 1
                                elif (108 - birthYear) > 80 and (108 - birthYear) <= 90:
                                    if '81-90歲' not in birthEraDict:
                                        birthEraDict['81-90歲'] = 0
                                        birthEraDict['81-90歲'] += 1
                                    else:
                                        birthEraDict['81-90歲'] += 1
                                elif (108 - birthYear) > 90:
                                    if '90歲以上' not in birthEraDict:
                                        birthEraDict['90歲以上'] = 0
                                        birthEraDict['90歲以上'] += 1
                                    else:
                                        birthEraDict['90歲以上'] += 1
                    print(birthEraDict)
                    hotelTenantTags = soup.select('.print a')
                    for url in hotelTenantTags:
                        url = url.get('href')
                        tenantDetailUrl = 'https://springtour.app{}'.format(url)
                        r = session_requests.get(tenantDetailUrl)
                        soup = bs(r.text, 'html.parser')
                        tenant = soup.select('#coupon_data .row p:last-of-type')
                        tenant = str(tenant[0]).replace('<p>', '').replace('人</p>', '')
                        hotelTenantTotal += int(tenant)
                        # print(tenant)
            print(birthPlaceDict)
            ws.append(['出生地', '人數'])
            for key, value in birthPlaceDict.items():
                ws.append([key, value])
            print(genderDict)
            ws.append(['', ''])
            ws.append(['性別', '人數'])
            for key, value in genderDict.items():
                ws.append([key, value])
            ws.append(['', ''])
            ws.append(['年齡層', '人數'])
            for key, value in birthEraDict.items():
                ws.append([key, value])

            ws2 = wb.create_sheet('統計表', 0)
            ws2.append(['出生地', '人數'])
            idDict = {'A': '臺北市', 'B': '臺中市', 'C': '基隆市', 'D': '臺南市', 'E': '高雄市', 'F': '新北市', 'G': '宜蘭縣', 'H': '桃園市', 'I': '嘉義市', 'J': '新竹縣', 'K': '苗栗縣', 'L': '臺中縣', 'M': '南投縣', 'N': '彰化縣', 'O': '新竹市', 'P': '雲林縣', 'Q': '嘉義縣', 'R': '台南縣', 'S': '高雄縣', 'T': '屏東縣', 'U': '花蓮縣', 'V': '臺東縣', 'W': '金門縣', 'X': '澎湖縣', 'Y': '陽明山管理局', 'Z': '連江縣', '1': '男', '2': '女'}
            idRowCount = 2
            for key, value in idDict.items():
                formula = '=IF(ISNA(VLOOKUP(A{},Sheet!$A:$B,2,FALSE)),0,VLOOKUP(A{},Sheet!$A:$B,2,FALSE))'.format(idRowCount, idRowCount)
                ws2.append([key, value, formula])
                idRowCount += 1
            formulaMale = '=IF(ISNA(VLOOKUP(A{},Sheet!$A:$B,2,FALSE)),0,VLOOKUP(A{},Sheet!$A:$B,2,FALSE))'.format(idRowCount, idRowCount)
            formulaFemale = '=IF(ISNA(VLOOKUP(A{},Sheet!$A:$B,2,FALSE)),0,VLOOKUP(A{},Sheet!$A:$B,2,FALSE))'.format(idRowCount, idRowCount)
            birthEraTitleList = ['10歲(含)以下', '11-20歲', '21-30歲', '31-40歲', '41-50歲', '51-60歲', '61-70歲', '71-80歲', '81-90歲', '90歲以上']
            for title in birthEraTitleList:
                formula = '=IF(ISNA(VLOOKUP(B{},Sheet!$A:$B,2,FALSE)),0,VLOOKUP(B{},Sheet!$A:$B,2,FALSE))'.format(idRowCount, idRowCount)
                idRowCount += 1
                ws2.append(['', title, formula])
            ws2.append(['', '住宿總人數', hotelTenantTotal])
            wb.save('{}/{} {} 旅客統計{}.xlsx'.format(folderName, hotelCount, hotelName, date))
            hotelCount += 1
            totalTetantCount += hotelTenantTotal
            print(totalTetantCount)
            print('==================== Sleep Start! ====================')
            time.sleep(3)
            # print('==================== Sleep End! ====================')
# convert end =========================================
else:
    print(r.status_code)