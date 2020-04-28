import requests, datetime
from bs4 import BeautifulSoup as bs
from openpyxl import Workbook
session_requests = requests.session()
loginUrl = 'https://springtour.app/Ajax/loginadmin'
payload = {
    'account': 'xxxxxx',
    'password': 'xxxxxx'
}
headers = {
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Origin': 'https://springtour.app',
    'Referer': 'https://springtour.app/Main/loginAdmin',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest'
}
getDateTime = str(datetime.datetime.now()).replace('-', '').replace(':', '')[0:13]
r = session_requests.post(loginUrl, data=payload, headers=headers)
if r.status_code == requests.codes.ok:
    print('Login Success!')
    wb = Workbook()
    ws = wb.active
    ws.append(['編號', '名稱', '資料筆數', '房間數', '房價合計', '補助合計'])
    count = 1
    for j in range(1, 33):
        targetUrl = 'https://springtour.app/Gov/statis/{}'.format(j)
        r = session_requests.get(targetUrl, headers = dict(referer = targetUrl))
        soup = bs(r.text, 'html.parser')
        td_tags = soup.find_all('td')
        i = 0
        info = []
        for td in td_tags:
            if i == 0:
                info.append(count)
                i += 1
                count += 1
            if i == 1:
                info.append(str(td.contents[0]))
                i += 1
            elif i == 2 or i < 6:
                info.append(int(td.contents[0]))
                i += 1
            elif i == 6:
                ws.append(info)
                i = 0
                info = []
        print('page {} success!'.format(j))
    wb.save('台中旅宿業者系統資料 {}.xlsx'.format(getDateTime))
else:
    print(r.status_code)