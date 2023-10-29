from flask import Flask, request, jsonify
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
from flask_cors import CORS
import random
import schedule
import threading
import re
from datetime import timedelta

app = Flask(__name__)
CORS(app)

# 設定 Google Sheets 權限範圍和 JSON 憑證文件路徑
scope = "https://www.googleapis.com/auth/spreadsheets " + \
        "https://www.googleapis.com/auth/drive"
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

# 授權並獲取 Google Sheets 物件
client = gspread.authorize(credentials)
spreadsheet = client.open('賦能港員工打卡記錄')

# 選擇工作表
worksheet = spreadsheet.get_worksheet(0)

# 函數用於觀察並處理下班打卡
def handle_end_of_day_attendance():
    tz = pytz.timezone('Asia/Taipei')
    current_time = datetime.now(tz)
    current_date = current_time.date()
    date_str = current_date.strftime('%Y/%m/%d')

    # 取得第一欄的所有值
    date_column = worksheet.col_values(1)
    # 現在日期
    current_date = datetime.now().date()
    # 格式化為與 Google Sheets 中日期格式相符的字符串
    date_str = current_date.strftime('%Y/%m/%d')
    # 定義正則表達式模式來匹配日期部分
    date_pattern = r'(\d{4}/\d{2}/\d{2})'

    # 遍歷日期欄位
    for row, date_value in enumerate(date_column, start=1):
        match = re.search(date_pattern, date_value)
        if match:
            if match.group(1) == date_str:
                員工姓名 = worksheet.cell(row, 2).value
                if 員工姓名 == "Tom":
                    continue

                if worksheet.cell(row, 3).value == "上班":
                    隨機秒數 = random.randint(0, 600)
                    打卡時間 = ((current_time - datetime(1899, 12, 30, tzinfo=tz)).total_seconds() + 隨機秒數 - 360) / (24 * 60 * 60)
                    員工姓名 = worksheet.cell(row, 2).value
                    worksheet.append_row([打卡時間, 員工姓名, "下班", "", "", "", "", ""])
                # 排序工作表。假設您的日期和時間位於第1列。
                worksheet.sort((1, 'asc'))

# 設定每天 18 點執行 handle_end_of_day_attendance 函數
schedule.every().day.at("15:24:10").do(handle_end_of_day_attendance)

# 啟動處理下班打卡的執行緒
def run_schedule():
    while True:
        schedule.run_pending()

@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    tz = pytz.timezone('Asia/Taipei')  # 以台灣為例
    current_time = datetime.now(tz)
    隨機秒數 = 0
    try:
        data = request.json
        
        if data is None:
            print("Received None from frontend")
            return jsonify({'error': 'Invalid JSON payload'})

        print("Received data from frontend:", data)

        # 從 JSON 資料中提取各種欄位
        員工姓名列表 = data.get('employeeName')  # 這現在是一個列表
        出缺勤狀況 = data.get('attendanceStatus')
        假別 = data.get('workOption')
        開始時間 = data.get('StartTime') or data.get('dateTimePicker1')
        結束時間 = data.get('EndTime') or data.get('dateTimePicker2')
        WFH原因 = data.get('WFHSection')
        
        if 員工姓名列表:
            員工姓名列表 = 員工姓名列表.split(',')

            # 然後遍歷員工姓名列表，為每一個員工添加一個新的行
            if 員工姓名列表:
                for 員工姓名 in 員工姓名列表:
                    員工姓名 = 員工姓名.strip()
                    隨機秒數 = random.randint(0, 600)
                    打卡時間 = ((current_time - datetime(1899, 12, 30, tzinfo=tz)).total_seconds() + 隨機秒數 - 360) / (24 * 60 * 60)
                    worksheet.append_row([打卡時間, 員工姓名, 出缺勤狀況, 假別, 開始時間, 結束時間, WFH原因])
            # 排序工作表。假設您的日期和時間位於第1列。
            worksheet.sort((1, 'asc'))

        print("Data written to Google Sheets successfully!")
        return jsonify({"message": "打卡資料已成功儲存"})
    
    except Exception as e:
        print("Error:", str(e))
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    # 啟動處理下班打卡的執行緒
    end_of_day_thread = threading.Thread(target=run_schedule)
    end_of_day_thread.daemon = True  # 執行緒隨著主程序退出而退出
    end_of_day_thread.start()

    app.run(port=5000)