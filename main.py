import requests
from lxml import html
import xlrd
import xlwt


def get_account():
    account = []
    wb = xlrd.open_workbook("data.xlsx")
    ws = wb.sheets()[0]
    for x in range(1, ws.nrows):
        account.append({
            'name': ws.cell_value(x, 0),
            'sfzh': ws.cell_value(x, 1),
            'mm': ws.cell_value(x, 2)}
        )
    return account


def cx(name, sfzh, mm):
    data = []
    session_requests = requests.session()
    result = session_requests.post(
        'http://wx.zzgjj.com/pcweb/pcchaxun/chaxun',
        data={'name': name, 'sfzh': sfzh, 'mm': mm},
        headers={'referer': 'http://wx.zzgjj.com/pcweb/pcchaxun/chaxun'}
    )
    tree = html.fromstring(result.content)
    elems = tree.findall(".//div[@class='cx']/")
    if elems:
        tmp_info = {}
        for elem in elems:
            info = elem.text.split('：')
            if info[0] == '公积金缴存信息':
                if tmp_info:
                    data.append(tmp_info)
                    tmp_info = {}
                continue
            tmp_info[info[0]] = info[1]
        data.append(tmp_info)
    else:
        data.append({'缴存人姓名': name, '缴存状态': '未开户或密码错误'})
    return data


def main():
    account = get_account()
    info = []
    for x in account:
        info += cx(x['name'], x['sfzh'], x['mm'])
    wb = xlwt.Workbook()
    ws = wb.add_sheet('查询结果', cell_overwrite_ok=True)
    ws_hd = ['公积金账户', '单位信息', '开户日期', '缴存人姓名', '缴存基数', '月缴额', '个人缴存比例', '单位缴存比例', '缴存余额', '缴至月份', '缴存状态']
    for x in range(len(info)):
        for y in range(len(ws_hd)):
            if x == 0:
                ws.write(x, y, ws_hd[y])
            ws.write(x + 1, y, info[x][ws_hd[y]])
    wb.save('result.xls')

if __name__ == '__main__':
    main()
    input()
