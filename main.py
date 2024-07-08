from selenium import webdriver
from chromedriver_py import binary_path
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import re
from bs4 import BeautifulSoup
import requests
import datetime
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import os
import json

# ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

# Variables - GitHub
line_notify_id = os.environ['LINE_NOTIFY_ID']
sheet_key = os.environ['GOOGLE_SHEETS_KEY']
gs_credentials = os.environ['GS_CREDENTIALS']

# Variables - Google Colab
# line_notify_id = LINE_NOTIFY_ID
# sheet_key = GOOGLE_SHEETS_KEY
# gs_credentials = GS_CREDENTIALS

# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

# LINE Notify ID
LINE_Notify_IDs = list(line_notify_id.split())

# 定義查找nid代碼函數
def find_nid(title, text):
    title_line_numbers = []
    for i, line in enumerate(text.split('\n')):
        if title in line:
            title_line_numbers.append(i)

    if not title_line_numbers:
        print(f'Cannot find "{title}" in the text.')
        return None

    title_line_number = title_line_numbers[0]
    title_line = text.split('\n')[title_line_number]

    nid_start_index = title_line.index('nid="') + 5
    nid_end_index = title_line.index('"', nid_start_index)
    nid = title_line[nid_start_index:nid_end_index]

    return nid

# 取得網頁內容
def get_content(url):
  # 發送GET請求獲取網頁內容
  response = requests.get(url)

  # 解析HTML內容
  soup = BeautifulSoup(response.content, 'html.parser')

  # 找到所有的 <p> 標籤
  p_tags = soup.find_all('p')

  # 整理文字內容
  text_list = []
  for p in p_tags:
      text = p.text.strip()
      text_list.append(text)
  text = ' '.join(text_list)
  text = ' '.join(text.split())  # 利用 split() 和 join() 將多個空白轉成單一空白
  # text = text.replace(' ', '\n')  # 將空白轉換成換行符號
  text = text.replace(' ', '')  # 刪除空白
  return text

text_limit = 1000-4-3

# LINE Notify
def LINE_Notify(school, category, date, title, unit, link, content):

  send_info_1 = f'【{school}】【{category}】{title}\n⦾公告日期：{date}\n⦾發佈單位：{unit}'
  send_info_2 = f'⦾內容：' if content != '' else ''
  send_info_3 = f'⦾更多資訊：{link}'

  text_len = len(send_info_1) + len(send_info_2) + len(send_info_3)
  if content != '':
    if text_len + len(content) > text_limit:
      content = f'{content[:(text_limit - text_len)]}⋯'
    params_message = f'{send_info_1}\n{send_info_2}{content}\n{send_info_3}'
  else:
    params_message = f'{send_info_1}\n{send_info_3}'

  for LINE_Notify_ID in LINE_Notify_IDs:
    headers = {
            'Authorization': 'Bearer ' + LINE_Notify_ID,
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    params = {'message': params_message}

    r = requests.post('https://notify-api.line.me/api/notify',
                            headers=headers, params=params)
    print(r.status_code)  #200

# Google Sheets 紀錄
scope = ['https://www.googleapis.com/auth/spreadsheets']
info = json.loads(gs_credentials)

creds = Credentials.from_service_account_info(info, scopes=scope)
gs = gspread.authorize(creds)

def google_sheets_refresh():

  global sheet, worksheet, rows_sheets, df

  # 使用表格的key打開表格
  sheet = gs.open_by_key(sheet_key)
  worksheet = sheet.get_worksheet(0)

  # 讀取所有行
  rows_sheets = worksheet.get_all_values()
  # 使用pandas創建數據框
  df = pd.DataFrame(rows_sheets)

def main(urls_temp):

    school, category, url = urls_temp.split('@')
    print(f'⦾{school} {category} {url}')

    # chromedriver 設定
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    service = Service(binary_path)
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(url)

    # 等待網頁載入完成
    driver.implicitly_wait(10)

    # 找到表格元素
    res = requests.get(url)
    soup = BeautifulSoup(res.text, 'html.parser')
    table_div = driver.find_element(By.ID, 'div_table_content')
    table = table_div.find_element(By.ID, 'ntb')
    html = table.get_attribute('outerHTML')

    # 解析HTML文件
    soup = BeautifulSoup(html, 'html.parser')

    # 格式化HTML文件
    formatted_html = soup.prettify()
    # print(formatted_html)

    # 找到表格中的所有資料列
    rows = table.find_elements(By.TAG_NAME, 'tr')

    # 打印每一行的 HTML 內容
    # for row in rows:
    #     row_html = row.get_attribute('outerHTML')
    #     print(row_html)

    # 定義需要查找的最新幾筆資料（最多9筆）
    numbers_of_new_data = 9

    # 印出最新幾筆資料的標題、單位和連結
    for i in range(numbers_of_new_data):
        row = rows[numbers_of_new_data - i]

        headers = []
        first_row = rows[0]
        header_cells = first_row.find_elements(By.TAG_NAME, 'th')
        for cell in header_cells:
          headers.append(cell.text)
          # print(cell.text)
        # print(headers)

        # row_html = row.get_attribute('outerHTML')
        # print(row_html)
        cells = row.find_elements(By.TAG_NAME, 'td')
        date = cells[headers.index('時間')].text if '時間' in headers else '-'
        title = cells[headers.index('標題')].text if '標題' in headers else '-'
        unit = cells[headers.index('單位')].text if '單位' in headers else '-'

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(row.get_attribute('outerHTML'), 'html.parser')

        # 找到 nid 的值
        nid = soup.find('tr')['nid']

        link_publish = f"{url[:url.find('ischool')]}ischool/public/news_view/show.php?nid={nid}"
        link = f"{url[:url.find('ischool')]}ischool/public/news_view/show.php?nid={nid}"
        content = get_content(link_publish)
        print(f'date:{date}\tcategory:{category}\ttitle:{title}\tunit:{unit}\tnid:{nid}\tlink:{link}\tcontent:{content}')

        # 獲取當前日期
        today = datetime.date.today()

        # 將日期格式化為2023/02/11的形式
        formatted_date = today.strftime("%Y/%m/%d")

        # 檢查nid是否已經存在於表格中
        sent = not(str(int(nid)) in nids)

        if sent:

          # 檢查標題是否已經存在於表格中
          titles = df[3].tolist()
          if title in titles:
            continue

          # 獲取新行
          now = datetime.datetime.now() + datetime.timedelta(hours=8)
          new_row = [now.strftime("%Y-%m-%d %H:%M:%S"), school, category, date, title, unit, nid, link, content]

          # 將新行添加到工作表中
          worksheet.append_row(new_row)

          # 獲取新行的索引
          new_row_index = len(rows) + 1

          # 更新單元格
          cell_list = worksheet.range('A{}:I{}'.format(new_row_index, new_row_index))
          for cell, value in zip(cell_list, new_row):
              cell.value = value
          worksheet.update_cells(cell_list)

          # 更新nids列表
          nids.append(int(nid))

          # 傳送至LINE Notify
          print(f'Sent: {nid}', end=' ')
          LINE_Notify(school, category, date, title, unit, link, content)

        # 刪除nid
        del nid

    # 關閉網頁
    driver.quit()

if __name__ == "__main__":

  # 開啟網頁
  urls = [
      '僑泰高級中學@即時訊息@https://www.ctas.tc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b53962254b4a2ce8333fe4d28a82127981050439&maximize=1&allbtn=0', 
      '僑泰高級中學@消息公佈欄@https://www.ctas.tc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_995edb94bf265937aca1457900a84c4673a6c773&maximize=1&allbtn=0', 
      '僑泰高級中學@首頁榮譽榜@https://www.ctas.tc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_c7b87337a13922e7868631cf15b9304e134bd194&maximize=1&allbtn=0', 
      '桃園市立內壢高級中等學校@最新公告@https://www.nlhs.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_426_2_5952ed0c12c4d36bd88cc820bec5a24e1ea088ba&maximize=1&allbtn=0', 
      '桃園市立內壢高級中等學校@考試專區@https://www.nlhs.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_426_2_9a8e9556c0482fa3ee97960b1e758bc40a32a6fd&maximize=1&allbtn=0', 
      '國立屏北高級中學@重要公告@https://www.ppsh.ptc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_100_2_7da0a53f3a6319dfe24c37cf48fac22f40678e27&maximize=1&allbtn=0', 
      '國立屏北高級中學@學生訊息@https://www.ppsh.ptc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_100_2_9f0f2175b7cbc7fe211ca6d27071c06172b0a69f&maximize=1&allbtn=0', 
      '國立虎尾高級中學@焦點看板@https://www.hwsh.ylc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_2cab01ac31360396f811bd99f531dd98e27b6e6c&maximize=1&allbtn=0', 
      '國立臺南大學附屬高級中學@學生公佈欄@https://www.tntcsh.tn.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b7f900216add60fd48b6d7aefc4ee317d7971c42&maximize=1&allbtn=0', 
      '臺南市黎明高級中學@校園活動@https://www.lmsh.tn.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_91b0f7a263d038f60ec87df0cd745419f792a223&maximize=1&allbtn=0', 
      '雲林縣義峰高級中學@最新消息@http://www.yfsh.ylc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_a2c4db1b778c971d21fe244011b3cc547850073f&maximize=1&allbtn=0', 
      '國立北港高級農工職業學校@消息公佈欄@https://www.pkvs.ylc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_ae7789b57835bad294fb23aedd9efa68d4540a18&maximize=1&allbtn=0', 
      '國立旗山農工@校內訊息@https://www.csvs.khc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_08ede4ae867c7fdd04cd0077a2c2fa60e9905749&maximize=1&allbtn=0', 
      '國立旗山農工@最新公告@https://www.csvs.khc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_041b051fce8fa7dd994067259f3bd91e4474c53f&maximize=1&allbtn=0', 
      '南投縣立旭光高級中學@校內活動公告@https://ischool.skjhs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f42d7a71ae901ef679dbeadeb09b06c5e90308aa&maximize=1&allbtn=0', 
      '新民高級中學@最新公告@https://www.shinmin.tc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b3af87feefe3f7a61583c5f6ceedcc978719aceb&maximize=1&allbtn=0', 
      '臺中市私立嶺東高級中學@最新消息@https://www.lths.tc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_e7226114a0083ffc6374db7e16e9438d122c83a4&maximize=1&allbtn=0', 
      '宜寧高級中學@消息公佈欄@https://www.inhs.tc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_ca65ff07cb6c0a2c371cb63a0960c53345dd8dc4&maximize=1&allbtn=0', 
      '國立竹北高級中學@綜合性公告@https://www.cpshs.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f92095519309688cd3d60355cbadcf5c5a8c97f4&maximize=1&allbtn=0', 
      '國立竹北高級中學@升學訊息@https://www.cpshs.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b20a054c748d5a6de895f09ff2053d911b70372c&maximize=1&allbtn=0', 
      '新北市私立格致高級中學@最新消息@http://www.gjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_858411e49cf3f3df5418ff9115a18f7a6a37f160&maximize=1&allbtn=0', 
      '天主教光仁高級中學@最新消息@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_0175a41dca498eab35e73c7c40fd1c141d1f3a58&maximize=1&allbtn=0', 
      '天主教光仁高級中學@榮譽榜@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_0f31e8a5a7bc3c4ef6609345e33cd3ae6b3e97cc&maximize=1&allbtn=0', 
      '天主教光仁高級中學@校務通報@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b1568f22c498d41d46e53b69325a31e78d51c87c&maximize=1&allbtn=0', 
      '天主教光仁高級中學@研習@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_e04b673ba655cf5fbdbab4e815d941418b3ec90c&maximize=1&allbtn=0', 
      '天主教光仁高級中學@師生活動與競賽@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_548130b3d109474559de5f5f564d0a729c3c7b3d&maximize=1&allbtn=0', 
      '天主教光仁高級中學@防疫專區@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_36ca124c0de704b62aead8039c9bbe125095f5aa&maximize=1&allbtn=0', 
      '天主教光仁高級中學@招標公告@https://www.kjsh.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_27c842d8808109d6838dde8f0c1222d6177d71b8&maximize=1&allbtn=0', 
      '清傳高商iSchool網站@消息公佈欄@https://www.ccvs.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_fad9d3fb5de4390b3dbc3b90f4f7e093afd1dcdd&maximize=1&allbtn=0', 
      '清傳高商iSchool網站@最新防疫資訊@https://www.ccvs.ntpc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_ae1e1e61bead297f93f57489a9834542090c4288&maximize=1&allbtn=0', 
      '桃園市私立光啟高級中學@消息公佈欄@http://www.phsh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_c7f4937b85e8e84f78d58a66519f811c86a6141c&maximize=1&allbtn=0', 
      '桃園市清華高級中學@最新消息@https://www.chvs.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_779a3305a97957a4ef3b815d3467fe81a166f6ef&maximize=1&allbtn=0', 
      '至善高級中學@校園公佈欄@https://www.lovejs.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_0069396de715eee9b91b2b10961009ac457a09db&maximize=1&allbtn=0', 
      '桃園市方曙商工高級中等學校@榮譽榜@https://www.fsvs.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_86ab5bcd710b3f8fadca88c962996282c8c6ec87&maximize=1&allbtn=0', 
      '桃園市立桃園高級中等學校@消息公佈欄@https://www.tysh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_74f1a2b604fce5e229dca49d8e696ba9792f838e&maximize=1&allbtn=0', 
      '桃園市立桃園高級中等學校@研習及競賽訊息@https://www.tysh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_89d12f2311dc06685bcfcbe87e249b9d5f097670&maximize=1&allbtn=0', 
      '桃園市立桃園高級中等學校@師生榮譽榜@https://www.tysh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_bb9f77fd0b92540eeb801f328cf417b17dc37952&maximize=1&allbtn=0', 
      '桃園市立大溪高級中等學校@校園公告@https://www.dssh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_c72ab8f583d816355b88e614f8e192e3156221ca&maximize=1&allbtn=0', 
      '桃園市立羅浮高級中等學校@最新消息@https://www.lfsh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_95ed0e4ef677eacf0282388e43d129bc189b303f&maximize=1&allbtn=0', 
      '桃園市立羅浮高級中等學校@媒體報導@https://www.lfsh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_cd7b3f721a7722a0339ea3004a05418ff2051e31&maximize=1&allbtn=0', 
      '桃園市立羅浮高級中等學校@競賽成績@https://www.lfsh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_e05438a436eb8163f26b9905b3e1179857137717&maximize=1&allbtn=0', 
      '桃園市立羅浮高級中等學校@學生活動@https://www.lfsh.tyc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_71203d01dd735f8f76e373852158ada3a1f96c6e&maximize=1&allbtn=0', 
      '國立竹北高級中學@綜合性公告@https://www.cpshs.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f92095519309688cd3d60355cbadcf5c5a8c97f4&maximize=1&allbtn=0', 
      '國立竹北高級中學@升學訊息@https://www.cpshs.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b20a054c748d5a6de895f09ff2053d911b70372c&maximize=1&allbtn=0', 
      '內思高工@榮譽榜公告@https://www.savs.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_ae6ed98182c98138c3adb6110d7f89a0846ec7e0&maximize=1&allbtn=0', 
      '內思高工@綜合性公告@https://www.savs.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_4bd3b2f2761a2796ff9b8c80366aa62f9e672853&maximize=1&allbtn=0', 
      '新竹縣東泰高級中學@消息公佈欄@https://www.ttsh.hcc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_c5079b8142faff65e3f1e35bce76e3910c4ed43a&maximize=1&allbtn=0', 
      '苗栗縣私立中興高級商工職業學校@教務處@https://www.csvs.mlc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_e9acbb7e760120f48b063179c7ecdf48aa3c4a99&maximize=1&allbtn=0', 
      '苗栗縣私立中興高級商工職業學校@最新消息@https://www.csvs.mlc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_a9ba948c2e244c6822f1df073c7e73319dbce6a9&maximize=1&allbtn=0', 
      '國立溪湖高級中學@榮譽榜@https://www.hhsh.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_8479e7d2e7ab201589e6af361db9695b9fb1ea3f&maximize=1&allbtn=0', 
      '國立溪湖高級中學@最新消息 News@https://www.hhsh.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_3ed70569e50c216bbfd52affdfa4f70a6e43e870&maximize=1&allbtn=0', 
      '國立員林高級農工職業學校@消息公佈欄@https://www.ylvs.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_3a697d6e56fcc365968005846b22476a8964f1d7&maximize=1&allbtn=0', 
      '國立員林高級農工職業學校@榮譽榜@https://www.ylvs.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_7d96a4ab2ed527faa4a920e09d6b0cb55740e46d&maximize=1&allbtn=0', 
      '國立員林高級農工職業學校@學生資訊@https://www.ylvs.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_d0c548cf05ba545ec93d160c7cdfb39283830c07&maximize=1&allbtn=0', 
      '國立員林高級農工職業學校@教師資訊@https://www.ylvs.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_b3262b016df9aac06287815e30d33d053e34a4fd&maximize=1&allbtn=0', 
      '國立員林家商@最新消息@https://www.ylhcvs.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_1303366a388017702188cda6f6b01923b445b86a&maximize=1&allbtn=0', 
      '彰化縣立和美高級中學@公告訊息@https://www.hmjh.chc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_07c6a54400b3c4642ab50a114a0a1c6949b5a79a&maximize=1&allbtn=0', 
      '教育部核定優質化國立埔里高工@校園公佈欄@https://www.plvs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_778d20bd132ae5d0cf0a7f35c3e79698c9289c95&maximize=1&allbtn=0', 
      '教育部核定優質化國立埔里高工@各科活動花絮及榮譽榜@https://www.plvs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_70f362cfedbe2ceb175b0377f4bd56b2beaab375&maximize=1&allbtn=0', 
      '教育部核定優質化國立埔里高工@師生競賽研習等訊息@https://www.plvs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_563e91ae749070f46463a5bcab817fee83e81f4e&maximize=1&allbtn=0', 
      '南投縣私立三育高級中學@消息公佈欄@https://www.taa.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f1b18cc02640f18edbe419d124ae9990999f6852&maximize=1&allbtn=0', 
      '均頭國中@消息公佈欄@https://www.jtjhs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_d1e5ed3e37b21654ee3cae94bcd41115cf47f02c&maximize=1&allbtn=0', 
      '均頭國中@榮譽榜@https://www.jtjhs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_1f0e913d80997c4d524675947a1395addfd2e27e&maximize=1&allbtn=0', 
      '均頭國中@校園徵才@https://www.jtjhs.ntct.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_df7129c3abc72ae41e6744067493d24b6637c92b&maximize=1&allbtn=0', 
      '國立後壁高級中學@最新消息@https://www.hpsh.tn.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_eeab16093903e04180cab07d430ace5d23b25200&maximize=1&allbtn=0', 
      '雲林縣私立大德工商職業學校@消息公佈欄@http://www.ddvs.ylc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_ce2683423bb4a0821f386ac093080c54cb4895ea&maximize=1&allbtn=0', 
      '雲林縣私立大德工商職業學校@榮譽榜@http://www.ddvs.ylc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f4ab2f04c0f58e62eb16442b6ae647eb6df8d0fa&maximize=1&allbtn=0', 
      '雲林縣私立大德工商職業學校@各處室公告@http://www.ddvs.ylc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_bbc0401ea3abeaa2716fcd57c137255bd74390b6&maximize=1&allbtn=0', 
      '國立後壁高級中學@最新消息@https://www.hpsh.tn.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_eeab16093903e04180cab07d430ace5d23b25200&maximize=1&allbtn=0', 
      '臺南市黎明高級中學@校園活動@https://www.lmsh.tn.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_91b0f7a263d038f60ec87df0cd745419f792a223&maximize=1&allbtn=0', 
      '國立旗美高級中學@本校訊息公告@https://www.cmsh.khc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_8a7d68a5b64a56c3777238a7b2470c8b167c37f1&maximize=1&allbtn=0', 
      '國立旗山農工@校內訊息@https://www.csvs.khc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_08ede4ae867c7fdd04cd0077a2c2fa60e9905749&maximize=1&allbtn=0', 
      '國立旗山農工@最新公告@https://www.csvs.khc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_041b051fce8fa7dd994067259f3bd91e4474c53f&maximize=1&allbtn=0', 
      '高雄市高苑工商職業學校@消息公佈欄@https://www.kyvs.kh.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_d81ae2ad38f0a64228a10b5323d7bde456cbb2f9&maximize=1&allbtn=0', 
      '國立花蓮高農@校園公佈欄@https://www.hla.hlc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_ee29ffd5c44f17089e50d56d06f4ede851e994a4&maximize=1&allbtn=0', 
      '國立花蓮高農@榮譽事蹟@https://www.hla.hlc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_56b79413f56a37009a049a61543f19585fdbdd08&maximize=1&allbtn=0', 
      '上騰工商@消息公佈欄@https://www.chvs.hlc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_fd31bea09509ec66b34925c3021a25f28c5dcbdf&maximize=1&allbtn=0', 
      '基隆光隆家商@消息公佈欄@https://www.klhcvs.kl.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_2afa5f0e53c7e09ac960abf79445796d39c5cafe&maximize=1&allbtn=0', 
      '國立新竹高級中學@最新消息@https://www.hchs.hc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_0516b5aba93b58b0547367faafb2f1dbe2ebba4c&maximize=1&allbtn=0', 
      '國立新竹高級工業職業學校@消息公佈欄@https://www.hcvs.hc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_9cde0cbcd39f55a57cb0d83c8a664cfd68a38895&maximize=1&allbtn=0', 
      '新竹世界高中@最新公告@https://www.wvs.hc.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_19ba14dad893b6c3939029848b0351ae9cf26912&maximize=1&allbtn=0', 
      '國立嘉義家職@最新消息@https://www.cyhvs.cy.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f9cb4dd227aa767e7494223ca89ccf8f44339512&maximize=1&allbtn=0', 
      '嘉義市天主教立仁高中@消息公佈欄@http://www.ligvs.cy.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_bef2a901c3271135b5377bcc1065df9722d24ee6&maximize=1&allbtn=0', 
      '臺北市協和祐德高級中等學校@校園公佈欄@https://www.hhvs.tp.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_404752329f61da8a7819e5c924b6c1f59fa2bc9f&maximize=1&allbtn=0', 
      '國立政治大學附屬高級中學@消息公佈欄@https://www.ahs.nccu.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_55ae02d68853299d86d2955bea3b3314a2651d4b&maximize=1&allbtn=0', 
      '國立政治大學附屬高級中學@活動訊息@https://www.ahs.nccu.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_f149ac6b4fe819faf022de7c56a3ded3f684967a&maximize=1&allbtn=0', 
      '高雄市立中正高級中學@消息公佈欄@https://www.cchs.kh.edu.tw/ischool/widget/site_news/main2.php?uid=WID_0_2_d858f68d210b3976cc6ba11477c1b6e273c98d3a&maximize=1&allbtn=0', 
  ]

  # 刷新Google Sheets表格
  google_sheets_refresh()

  # 取得Google Sheets nids列表
  _nids = df[6].tolist()
  nids = []
  for n in _nids:
    try:
      nids.append(str(int(n)))
    except:
      continue

  for urls_temp in urls:

    finished = False
    try_times_limit = 2
    for _ in range(try_times_limit):
      try:
        main(urls_temp)
        finished = True
        break
      except:
        print('retrying...')
        next
    
    if not finished: print(f'error : {urls_temp}')
