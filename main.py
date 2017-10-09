import requests
from lxml import html
import xlrd
import xlwt


def get_data():
    data = []
    wb = xlrd.open_workbook("data.xlsx")
    ws = wb.sheets()[0]
    for x in range(1, ws.nrows):
        data.append({
            'name': ws.cell_value(x, 0),
            'sfzh': ws.cell_value(x, 1),
            'mm': ws.cell_value(x, 2)}
        )
    return data


def main():
    data = get_data()

    global row
    row = 0
    def cx(name, sfzh, mm):
        session_requests = requests.session()
        result = session_requests.post(
            'http://wx.zzgjj.com/pcweb/pcchaxun/chaxun',
            data = {'name': name, 'sfzh': sfzh, 'mm': mm},
            headers = {'referer': 'http://wx.zzgjj.com/pcweb/pcchaxun/chaxun'}
        )
        tree = html.fromstring(result.content)
        elems = tree.findall(".//div[@class='cx']/")
        if elems:
            for elem in elems:
                info = elem.text.split('：')
                # TODO: 改成数组
                print(info)
                global row
                if info[0] == '公积金缴存信息':
                    row += 1
                if info[0] == '公积金账户':
                    ws.write(row, 0, info[1])
                if info[0] == '单位信息':
                    ws.write(row, 1, info[1])
                if info[0] == '开户日期':
                    ws.write(row, 2, info[1])
                if info[0] == '缴存人姓名':
                    ws.write(row, 3, info[1])
                if info[0] == '缴存基数':
                    ws.write(row, 4, info[1])
                if info[0] == '月缴额':
                    ws.write(row, 5, info[1])
                if info[0] == '个人缴存比例':
                    ws.write(row, 6, info[1])
                if info[0] == '单位缴存比例':
                    ws.write(row, 7, info[1])
                if info[0] == '缴存余额':
                    ws.write(row, 8, info[1])
                if info[0] == '缴至月份':
                    ws.write(row, 9, info[1])
                if info[0] == '缴存状态':
                    ws.write(row, 10, info[1])
        else:
            row += 1
            ws.write(row, 3, name)
            ws.write(row, 10, '未开户或密码错误')


    wb = xlwt.Workbook()
    ws = wb.add_sheet('ws', cell_overwrite_ok=True)
    ws_hd = ['公积金账户', '单位信息', '开户日期', '缴存人姓名', '缴存基数',
        '月缴额', '个人缴存比例', '单位缴存比例', '缴存余额', '缴至月份', '缴存状态']
    for x in range(0, len(ws_hd)):
        ws.write(row, x, ws_hd[x])
    for x in data:
        cx(x['name'], x['sfzh'], x['mm'])
    wb.save('result.xls')

if __name__ == '__main__':
    main()
    input()
