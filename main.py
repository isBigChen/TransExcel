import time
from threading import Thread
import pandas as pd
import requests
import random
from hashlib import md5
import sys
from PySide6 import QtCore, QtWidgets
from PySide6.QtWidgets import QFileDialog
import untitled


class TransExcel(QtWidgets.QMainWindow, untitled.Ui_MainWindow):
    progress_signal = QtCore.Signal(int)

    def __init__(self, MainWindow):
        super(TransExcel, self).__init__()
        super().setupUi(MainWindow)

        self.progress_signal.connect(self.set_progress)
        self.progress_signal.emit(0)

        self.appid = ""
        self.appkey = ""
        self.from_file = ""
        self.to_file = ""

        self.get_info()

        self.pushButton.clicked.connect(self.select_file)
        self.pushButton_2.clicked.connect(self.trans_excel)

    def set_progress(self, val):
        self.progressBar.setValue(val)

    def get_info(self):
        try:
            with open('./ini.txt', 'r') as fp:
                data = fp.readlines()
                self.appid = data[0].split("=")[1].strip()
                self.appkey = data[1].split('=')[1].strip()
            fp.close()
        except Exception as e:
            self.textBrowser_2.append(e)

    def select_file(self):
        try:
            file_name = QFileDialog.getOpenFileName(self, "选择文件")
            self.from_file = file_name[0]
            self.to_file = self.from_file[:-5] + "_2.xlsx"
            self.textBrowser.setText(self.from_file)
        except Exception as e:
            self.textBrowser_2.append(e)

    def trans_excel(self):
        # operate_excel(self.from_file, 'Sheet1', self.to_file, 'Sheet1', self.appid, self.appkey)
        try:
            work = Thread(target=self.operate_excel,
                          args=(self.from_file, 'Sheet1', self.to_file, 'Sheet1', self.appid, self.appkey))
            work.start()
        except Exception as e:
            self.textBrowser_2.append(e)

    def operate_excel(self, from_file, from_sheet, to_file, to_sheet, appid, appkey):
        df = pd.read_excel(from_file, sheet_name=from_sheet)

        row_count = df.shape[0]
        col_count = df.shape[1]
        # print(row_count, col_count)

        for row_index in range(row_count):
            for col_index in range(col_count):
                src_data = str(df.iloc[row_index, col_index])
                # print(appid, appkey)
                df.iloc[row_index, col_index] = get_transfer_content(src_data, appid, appkey)
                self.progress_signal.emit(((row_index+1)/row_count)*100)
                time.sleep(1)

        df.to_excel(to_file, to_sheet, index=False)


def make_md5(s, encoding='utf-8'):
    return md5(s.encode(encoding)).hexdigest()


def get_transfer_content(query, appid, appkey, from_lang='zh', to_lang='kor'):
    endpoint = 'http://api.fanyi.baidu.com'
    path = '/api/trans/vip/translate'
    url = endpoint + path

    salt = random.randint(32768, 65536)
    sign = make_md5(appid + query + str(salt) + appkey)

    # Build request
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    payload = {'appid': appid, 'q': query, 'from': from_lang, 'to': to_lang, 'salt': salt, 'sign': sign}

    # Send request
    r = requests.post(url, params=payload, headers=headers)
    result = r.json()

    # Show response
    # print(json.dumps(result, indent=4, ensure_ascii=False))
    return str(result['trans_result'][0]['dst'])


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    QtWidgets.QApplication.setStyle(QtWidgets.QStyleFactory.create('Fusion'))
    mainWindow = QtWidgets.QMainWindow()
    ui = TransExcel(mainWindow)
    mainWindow.show()
    sys.exit(app.exec())

    # print(get_transfer_content("你好", "20230528001692064", "sYkfOSAGiNY5LA4T5atg", from_lang='zh', to_lang='kor'))
