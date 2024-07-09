from flask import Flask, render_template, request, send_file
import urllib.request
import datetime, re, json, os
from openpyxl import Workbook
from html import unescape
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

app = Flask(__name__)


#네이버 api
client_id = 'OsJDydLq5gAsxl6E00J0'
client_secret = 'JYQ4XeX1Vx'

#스케쥴러
scheduler = BackgroundScheduler()
scheduler.start()

def remove_html_tags(text):
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

def getRequestUrl(url):
    req = urllib.request.Request(url)
    req.add_header("X-Naver-Client-Id", client_id)
    req.add_header("X-Naver-Client-Secret", client_secret)
    
    try:
        response = urllib.request.urlopen(req)
        if response.getcode() == 200:
            print("[%s] Url Request Success" % datetime.datetime.now())
            return response.read().decode('utf-8')
    except Exception as e:
        print(e)
        print("[%s] Error for URL : %s" % (datetime.datetime.now(), url))
        return None

def getNaverSearch(node, srcText, start, display):
    base = "https://openapi.naver.com/v1/search"
    node = "/%s.json" % node
    parameters = "?query=%s&start=%s&display=%s" % (urllib.parse.quote(srcText), start, display)

    url = base + node + parameters
    responseDecode = getRequestUrl(url)

    if responseDecode is None:
        return None
    else:
        return json.loads(responseDecode)

def getPostData(post, jsonResult, cnt):
    title = unescape(remove_html_tags(post['title']))
    description = unescape(remove_html_tags(post['description']))
    org_link = unescape(remove_html_tags(post['originallink']))
    link = unescape(remove_html_tags(post['link']))

    # pubDate 형식을 datetime 객체로 변환
    try:
        pDate = datetime.datetime.strptime(post['pubDate'], '%a, %d %b %Y %H:%M:%S +0900')
        pDate = pDate.strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        pDate = None

    jsonResult.append({'cnt': cnt, 'title': title, 'description': description,
                       'org_link': org_link, 'link': link, 'pDate': pDate})
    return

def send_excel_via_email(filename, recipient):
    
    load_dotenv()
    SECRET_ID = os.getenv("SECRET_ID")
    SECRET_PASS = os.getenv("SECRET_PASS")
    msg = MIMEMultipart()
    msg['From'] = SECRET_ID
    msg['To'] = recipient
    msg['Subject'] = 'Scheduled Excel File'

    with open(filename, 'rb') as attachment:
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(filename)}')
        msg.attach(part)

    server = smtplib.SMTP('smtp.naver.com', 587)
    server.starttls()
    server.login(SECRET_ID, SECRET_PASS)
    server.sendmail(SECRET_ID, recipient, msg.as_string())
    server.quit()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        srcText = request.form['keyword']
        schedule_hour = int(request.form['schedule_hour'])
        schedule_minute = int(request.form['schedule_minute'])
        recipient = request.form['recipient']
        node = 'news'  # 크롤링할 대상
        cnt = 0
        jsonResult = []

        jsonResponse = getNaverSearch(node, srcText, 1, 100)
        total = jsonResponse['total']

        while jsonResponse and jsonResponse['display'] != 0:
            for post in jsonResponse['items']:
                cnt += 1
                getPostData(post, jsonResult, cnt)

            start = jsonResponse['start'] + jsonResponse['display']
            jsonResponse = getNaverSearch(node, srcText, start, 100)

        print('전체 검색 : %d 건' % total)
        print(f'받을 이메일주소: {recipient}')
        current_datetime = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        excel_filename = f'Subjects-subscribed-{node}.xlsx'
        file_path = os.path.join('scheduled_files', excel_filename)

        if not os.path.exists('scheduled_files'):
            os.makedirs('scheduled_files')

        wb = Workbook()
        ws = wb.active
        ws.append(['번호', '제목', '설명', '원본 링크', '링크', '날짜'])

        for item in jsonResult:
            ws.append([item['cnt'], item['title'], item['description'], item['org_link'], item['link'], item['pDate']])

        wb.save(file_path)

        # Schedule the task
        scheduler.add_job(
            send_excel_via_email, 'cron', hour=schedule_hour, minute=schedule_minute, args=[file_path, recipient]
        )

        return render_template('result.html', items=jsonResult, total=total, excel_filename=excel_filename)

    return render_template('index.html')

@app.route('/download/<path:filename>')
def download(filename):
    return send_file(os.path.join('scheduled_files', filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
