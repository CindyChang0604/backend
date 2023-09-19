from flask import Flask, request, jsonify
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
from flask_cors import CORS

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

@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    try:
        data = request.json
        
        if data is None:
            print("Received None from frontend")
            return jsonify({'error': 'Invalid JSON payload'})

        print("Received data from frontend:", data)

        # 從 JSON 資料中提取各種欄位
        tz = pytz.timezone('Asia/Taipei')  # 以台灣為例
        打卡時間 = datetime.now(tz)
        員工姓名列表 = data.get('employeeName')  # 這現在是一個列表
        出缺勤狀況 = data.get('attendanceStatus')
        假別 = data.get('workOption')
        開始時間 = data.get('StartTime') or data.get('dateTimePicker1')
        結束時間 = data.get('EndTime') or data.get('dateTimePicker2')
        WFH原因 = data.get('WFHSection')
        
        # 根據時間是上午還是下午添加上午/下午，並將時間格式化
        if 打卡時間.hour < 12:
            formatted_time = 打卡時間.strftime('%Y/%m/%d 上午 %I:%M:%S')
        else:
            formatted_time = 打卡時間.strftime('%Y/%m/%d 下午 %I:%M:%S')
        
        
        if 員工姓名列表:
            員工姓名列表 = 員工姓名列表.split(',')

            # 然後遍歷員工姓名列表，為每一個員工添加一個新的行
            if 員工姓名列表:
                for 員工姓名 in 員工姓名列表:
                    員工姓名 = 員工姓名.strip()
                    worksheet.append_row([打卡時間, 員工姓名, 出缺勤狀況, 假別, 開始時間, 結束時間, WFH原因])
        
        print("Data written to Google Sheets successfully!")
        return jsonify({"message": "打卡資料已成功儲存"})
    
    except Exception as e:
        print("Error:", str(e))
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(port=5000)