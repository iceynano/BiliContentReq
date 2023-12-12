import requests
import json
import xlwt
import re
from time import sleep
book = xlwt.Workbook()
sheet = book.add_sheet('评论', cell_overwrite_ok=True)
row0 = ['username', 'sex', 'level', 'like', 'time', 'comment']
row = 1
for i in range(0, len(row0)):
    sheet.write(0, i, row0[i])

def unzip(replies):
    global row
    if replies == None:
        return
    for rep in replies:
        content = rep['content']
        like = rep['like']
        msg = content['message']
        member = rep['member']
        sex = member['sex']
        username = member['uname']
        level_info = member['level_info']
        level = level_info['current_level']  # type int
        reply_control = rep['reply_control']
        time_desc = reply_control['time_desc']
        lis = [username, sex, level, like, time_desc, msg]
        for j in range(0, 6):
            sheet.write(row, j, lis[j])
        row += 1
        subreplies = rep['replies']
        if subreplies != None:
            unzip(subreplies)
            row += 1

def parser_comment(url, oid, pages, title):
    global row
    for page in range(1, pages):
        print(f'processing pages : {page}')
        url = f"https://api.bilibili.com/x/v2/reply?jsonp=jsonp&pn={page}&type=1&oid={oid}&sort=2"
        sleep(1)
        head = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36'
        }
        data = requests.get(url, headers=head).json()
        # try:
        data = data['data']
        replies = data['replies']
        unzip(replies)
        # except Exception as e:
        #     print(e)
        #     print("网络错误或解析错误， 请检查data文件")

url_list = []
cnt = 0

if __name__ == "__main__":
    print("请输入BV号：")
    st_r = input()
    strs = f"https://www.bilibili.com/video/{st_r}"
    aid_ = f"https://api.bilibili.com/x/web-interface/view?bvid={st_r}"
    head = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36'
    }
    tmp_data = requests.get(aid_, headers=head).json()
    print(f"正在从 {aid_} 解析网址")
    title = tmp_data['data']['title']
    aid = tmp_data['data']['aid']
    print("请输入要保存的结果文件名：")
    x = input()
    url_list.append(strs)
    for url in url_list:
        cnt += 1
        parser_comment(url, aid, 7, title)
    book.save(filename_or_stream=f"F:\\{x}.xls")