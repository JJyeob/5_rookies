from flask import Flask, render_template, request, send_file, jsonify
import urllib.request
import datetime, re, json, os, smtplib
from openpyxl import Workbook
from html import unescape
from apscheduler.schedulers.background import BackgroundScheduler
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

app = Flask(__name__)

# 네이버 API
client_id = 'OsJDydLq5gAsxl6E00J0'
client_secret = 'JYQ4XeX1Vx'

# 스케줄러 초기화 및 시작
scheduler = BackgroundScheduler()
scheduler.start()

# HTML 태그 제거 함수
def remove_html_tags(text):
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

# 네이버 API 요청 함수
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

# 네이버 검색 요청 함수
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

# 검색 결과에서 필요한 데이터 추출 함수
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

# 검색결과 가져오기
def nav_search_result(node, srcText):
    jsonResult = []
    cnt = 0

    # 네이버 검색 요청
    jsonResponse = getNaverSearch(node, srcText, 1, 100)
    total = jsonResponse['total']

    # 검색 결과 처리
    while jsonResponse and jsonResponse['display'] != 0:
        for post in jsonResponse['items']:
            cnt += 1
            getPostData(post, jsonResult, cnt)

        start = jsonResponse['start'] + jsonResponse['display']
        jsonResponse = getNaverSearch(node, srcText, start, 100)

    return jsonResult, total

#엑셀 파일생성
def create_excel_file(jsonResult, node):
    current_datetime = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    excel_filename = f'Naver_Subscribed_{node}_{current_datetime}.xlsx'
    file_path = os.path.join('scheduled_files', excel_filename)

    # 파일 저장 디렉토리 생성
    if not os.path.exists('scheduled_files'):
        os.makedirs('scheduled_files')

    # 엑셀 파일 생성 및 데이터 추가
    wb = Workbook()
    ws = wb.active
    ws.append(['번호', '제목', '설명', '원본 링크', '링크', '날짜'])

    for item in jsonResult:
        ws.append([item['cnt'], item['title'], item['description'], item['org_link'], item['link'], item['pDate']])

    wb.save(file_path)

    return file_path, excel_filename

#이메일 첨부
def send_email_excel(recipient, file_path, excel_filename,):
    load_dotenv()
    SECRET_ID = os.getenv("SECRET_ID")
    SECRET_PASS = os.getenv("SECRET_PASS")
    msg = MIMEMultipart()
    msg['From'] = SECRET_ID
    msg['To'] = recipient
    msg['Subject'] = '키워드 관련 기사'

    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
        msg.attach(part)

    server = smtplib.SMTP('smtp.naver.com', 587)
    server.starttls()
    server.login(SECRET_ID, SECRET_PASS)
    server.sendmail(SECRET_ID, recipient, msg.as_string())
    server.quit()

# 전체 작업수행함수(크롤링, 엑셀생성, 이메일전송)
def create_and_send_excel(srcText, recipient, node):
    jsonResult, total = nav_search_result(node, srcText)
    file_path, excel_filename = create_excel_file(jsonResult, node)
    send_email_excel(recipient, file_path, excel_filename)
    return jsonResult, excel_filename, total

@app.route('/', methods=['GET', 'POST'])
def index():
    # 메인 페이지로, 검색어와 이메일 주소를 입력받아 스케줄러에 작업을 추가하는 함수
    if request.method == 'POST':
        srcText = request.form['keyword']
        recipient = request.form['recipient']
        schedule_hour = int(request.form['schedule_hour'])
        schedule_minute = int(request.form['schedule_minute'])
        node = 'news' # 크롤링할 대상

        # 스케줄러 작업 설정 (1분마다 실행되도록 설정)
        # scheduler.add_job(
        # create_and_send_excel, 'interval', minutes=1, args=[srcText, recipient, node]
        # )
        scheduler.add_job(
        create_and_send_excel, 'cron', hour=schedule_hour, minute=schedule_minute, args=[srcText, recipient, node]
        )
        return render_template('result.html', keyword=srcText, recipient=recipient)

    return render_template('index.html')


@app.route('/download/<path:filename>')
def download(filename):
    return send_file(os.path.join('scheduled_files', filename), as_attachment=True)

@app.route('/get_latest_data', methods=['POST'])
def get_latest_data():
    srcText = request.form['keyword']
    recipient = request.form['recipient']
    node = 'news'

    # 최신 데이터를 가져오고 엑셀 파일을 생성
    jsonResult, excel_filename, total = create_and_send_excel(srcText, recipient, node)

    return jsonify({'items': jsonResult, 'total': total, 'excel_filename': excel_filename})

if __name__ == '__main__':
    app.run(debug=True)