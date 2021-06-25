# encoding: utf-8

from PyQt5.QtCore import Qt, QDateTime, QDate, QThread, QObject, pyqtSignal, QMutex
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QMessageBox, QFileDialog,
              QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QDateTimeEdit, 
              QTextEdit, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView)
import os
import sys
import traceback
import requests
from datetime import date, datetime, timedelta
from bs4 import BeautifulSoup as BS
from threading import Thread
from time import sleep
from queue import Queue
import cgitb
cgitb.enable(format = 'text')

AppVersion = "v0.1.1"

HttpHeaders = {
  "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.101 Safari/537.36"
}

RootURI   = 'http://www.ithome.com.tw'
TargetURI = "http://www.ithome.com.tw/taxonomy/term/5971/all"
PageParam = 'page'

ContentClassPost = '.item'
ContentClassDate = '.post-at'
ContentClassTitle = '.title'
SigWorking  = 0
SigComplete = 1
SigError    = 2
SigAbort    = 3

LastError  = None
MainWindow = None
MainApp    = None
Signal     = 0
Mutex      = QMutex()

FastSelectOptions = {
  '近七天': 7,
  '這個月': 'ThisMonth',
  '近 30 日': 30,
  '近半年': 182.621099,
  '今年': 'ThisYear',
  '去年': 'LastYear'
}

def get_error_logname():
  directory = '.'
  try:
    if not os.path.isdir('錯誤日誌'):
      os.mkdir('錯誤日誌')
    directory = '錯誤日誌'
  except Exception:
    return None
  ret = f"{directory}/errorlog_{str(datetime.now()).split('.')[0]}.log"
  _tr = str.maketrans({
    ':': '-', ' ': '_'
  })
  return ret.translate(_tr)

def safe_execute_func(func, args=[], kwargs={}):
  global MainWindow
  try:
    if args is None:
      args = []
    if kwargs is None:
      kwargs = {}
    return func(*args, **kwargs)
  except Exception as err:
    print("An error occurred:", err)
    err_info = traceback.format_exc()
    print(err_info)
    handle_exception(err, err_info, MainWindow)
  return None

def handle_exception(err, errinfo, window=None):
  dmp_ok = dump_errorlog(err, errinfo)
  if window:
    QMessageBox.critical(None, 'Error', '運行過程中產生錯誤, 詳情請見資訊窗格')
    if dmp_ok:
      window.append_log(' ')
      window.append_log(f"運行過程中產生錯誤, 請將紀錄檔 `{dmp_ok}` 寄送給開發人員以排除錯誤")
      window.append_log(f"錯誤內容: {err}")
      window.append_log("同時請確認您的輸出路徑無誤")
    else:
      window.append_log(f"運行過程中產生錯誤, 請將以下內容複製並寄送給開發人員以排除錯誤")
      window.append_log(str(err))
      window.append_log(errinfo)
      window.append_log("同時請確認您的輸出路徑無誤")
  else:
    if dmp_ok:
      QMessageBox.critical(None, 'Error', f'運行過程中產生錯誤, 請將紀錄檔 `{dmp_ok}` 寄送給開發人員以排除錯誤, 同時請確認您的檔案及內容是正確的', QMessageBox.Ok)
    else:
      _msg  = '運行過程中產生錯誤, 請將以下內容截圖並寄送給開發人員以排除錯誤, 同時請確認您的檔案及內容是正確的\n'
      _msg += str(err) + "\n" + errinfo
      QMessageBox.critical(None, 'Error', _msg, QMessageBox.Ok)

def dump_errorlog(err, errinfo):
  try:
    filename = get_error_logname()
    if not filename:
      return False
    with open(filename, 'w') as fp:
      fp.write(f"{str(err)}\n{errinfo}")
    return filename
  except Exception:
    return False

def str2date(ss, key='-'):
  try:
    return datetime(*[int(dat) for dat in ss.split(key)])
  except Exception:
    return None

def is_file_writable(filename):
    try:
      with open(filename, 'a') as _:
        pass
    except Exception:
      return False
    return True

def open_external_file(path):
  try:
    if sys.platform.startswith('linux'):
      os.system(f"xdg-open \"{path}\"")  
    elif sys.platform == 'win32':
      os.system(f"start /b \"\" \"{path}\"")
    else:
      os.system(f"start \"{path}\"")  
    return True
  except Exception:
    return False

### Main curl process
def is_internet_available():
  try:
    if requests.get('http://www.google.com', timeout=3):
      return True
  except Exception:
    pass
  return False

def parse_headline(doc):
  ret = []
  for post in list(doc.select(ContentClassPost)):
    obj = {}
    try:
      node = post.select(ContentClassTitle)[0]
      obj['title'] = node.text
      obj['link']  = f"{RootURI}{node.select('a')[0]['href']}"
    except Exception as err:
      print("An error occurred:", err)
      pass
    # date
    try:
      obj['date'] = post.select(ContentClassDate)[0].text.strip()
    except Exception as err:
      print("An error occurred:", err)
      pass
    if obj:
      ret.append(obj)
  return ret

def parse_content(doc):
  ret = []
  obj = {'tags': [], 'title': '', 'link': ''}
  for node in doc.select('strong'):
    line = node.text.strip()
    try:
      if (line[0] == '#' or line[0] == '＃'):
        obj['tags'] = line.split()
      elif line[0] == '詳' or '全文' in line:
        obj['link'] = node.select('a')[0]['href']
        print(obj)
        ret.append(obj)
        obj = {'tags': [], 'title': '', 'link': ''}
      else:
        obj['title'] = line
    except Exception as err:
      print(f"[WARNING] Unable to parse section `{line}` due to {err}")
  return ret

class QtCurlWorker(QObject):
  finished = pyqtSignal()
  errored  = pyqtSignal()
  aborted  = pyqtSignal()
  
  def __init__(self, parent_window, start_date, end_date):
    super(QtCurlWorker, self).__init__()
    self.parent_window = parent_window
    self.save_path = self.parent_window.save_path
    self.start_qdate = start_date
    self.end_qdate   = end_date

  def run(self):
    self.start_async(self.save_path)

  def verify_internet(self):
    self.parent_window.append_log("\n確認連線狀態...")
    if not is_internet_available():
      self.aborted.emit()
      self.parent_window.append_log("無網路連線")
      QMessageBox.critical(None, 'Error', '沒有網際網路連線!')
      return False
    else:
      self.parent_window.append_log("已連線")
    return True

  def start_async(self, output_path):
    global LastError, Signal
    try:
      self._start_async_proc(output_path)
      self.finished.emit()
    except Exception as err:
      print("An error occurred:", err)
      err_info = traceback.format_exc()
      LastError = (err, err_info)
      self.errored.emit()

  def _start_async_proc(self, output_path):
    if not self.verify_internet():
      return
    qdate_st = self.start_qdate
    qdate_ed = self.end_qdate
    date_st  = datetime(qdate_st.year(), qdate_st.month(), qdate_st.day())
    date_ed  = datetime(qdate_ed.year(), qdate_ed.month(), qdate_ed.day()) + timedelta(1)
    if date_ed < date_st:
      date_st,date_ed = date_ed,date_st
    page = 0
    html = None
    flag_outdated = False
    print(date_st, date_ed)
    try:
      file = open(output_path, 'w', encoding='utf8')
      file.write("標題,網址,分類\n")
      while not flag_outdated:
        # Index processing
        uri = f"{TargetURI}?{PageParam}={page}"
        print("Connecting to", uri)
        self.parent_window.append_log(f"正在連線至 {uri}")
        try:
          html = requests.get(uri, headers=HttpHeaders).content
          html = BS(html, 'html.parser')
        except Exception as err:
          print("An error occurred:", err)
          self.aborted.emit()
          self.parent_window.append_log(f"程式無法連線至 {uri}")
          QMessageBox.critical(None, 'Error', f"程式無法連線至 {uri}; 請檢查網路連線正常且該網站有上線.")
          break
        posts = parse_headline(html)
        page += 1
        print("Process page", page)
        # last page (has nothing)
        if not posts:
          break
        for post in posts:
          title = post['title'].split('：')[0]
          date  = str2date(post['date'])
          if date > date_ed:
            pass
          if date < date_st:
            flag_outdated = True
            break
          print(title, date)
          file.write(f"{title}, {post['link']}, \n")
          
          # get link to news post and extract tags
          self.parent_window.append_log(f"正在取得新聞內容 (標題: {title})")
          print(f"Getting content of {post['link']}")
          try:
            sleep(0.5)
            html = requests.get(post['link'], headers=HttpHeaders)
            html = BS(html.content, 'html.parser')
          except Exception as err:
            print("An error occurred:", err)
            self.aborted.emit()
            self.parent_window.append_log(f"程式無法連線至 {post['link']}")
            QMessageBox.critical(None, 'Error', f"程式無法連線至 {post['link']}; 請檢查網路連線正常且該網站有上線.")
            break
          print("Parsing page")
          self.parent_window.append_log(f"正在提取文章...")
          sections = parse_content(html)
          for section in sections:
            file.write(section['title']+', ')
            file.write(section['link']+', ')
            print(" ".join(section['tags']))
            file.write(" ".join(section['tags']) + '\n')
        # for each post
        sleep(1) # sleep a while to prevent excessive requests crash
      # while not outdated
    finally:
      print("Done")
      file.close()

### GUI section
class MainGUI(QMainWindow):
  request_logging = pyqtSignal()

  def __init__(self):
    super().__init__()
    self.resize(600, 400)
    self.setWindowTitle("iThome 資安周報爬蟲工具")
    self.setWindowOpacity(1)
    self.init_ui()
    self.setup_helper_window()
    self.log_queue = Queue()
    self.request_logging.connect(self._logging)
  
  def init_ui(self):
    self.main_widget = QWidget()  
    self.main_layout = QGridLayout()
    self.main_widget.setLayout(self.main_layout)
    self.setCentralWidget(self.main_widget) 

    self.main_layout.addWidget(QLabel('開始日期 :'), 0, 0, 1, 1)
    self.in_startdate = QDateTimeEdit(QDateTime.currentDateTime())
    self.in_startdate.setCalendarPopup(True)
    self.in_startdate.setDisplayFormat('yyyy-MM-dd')
    self.main_layout.addWidget(self.in_startdate, 0, 1, 1, 3)
    
    self.main_layout.addWidget(QLabel('結束日期 :'), 0, 4, 1, 1)
    self.in_enddate = QDateTimeEdit(QDateTime.currentDateTime())
    self.in_enddate.setCalendarPopup(True)
    self.in_enddate.setDisplayFormat('yyyy-MM-dd')
    self.main_layout.addWidget(self.in_enddate, 0, 5, 1, 3)

    self.main_layout.addWidget(QLabel('快速選擇 :'), 0, 8, 1, 1)
    self.cmb_fast_select = QComboBox()
    self.cmb_fast_select.addItems(list(FastSelectOptions.keys()))
    self.main_layout.addWidget(self.cmb_fast_select, 0, 9, 1, 2)
    self.btn_fast_select = QPushButton('選擇')
    self.btn_fast_select.clicked.connect(self.on_fast_select)
    self.main_layout.addWidget(self.btn_fast_select, 0, 11, 1, 1)


    self.edit_log = QTextEdit(f'初始化成功; 目標網址: {TargetURI}')
    self.edit_log.setReadOnly(True)
    self.main_layout.addWidget(self.edit_log, 1, 0, 4, 12)
    self.edit_log.moveCursor(QTextCursor.End)
    
    self.btn_save = QPushButton('選擇儲存路徑')
    self.btn_save.clicked.connect(lambda: safe_execute_func(self.execute_mainproc))
    self.main_layout.addWidget(self.btn_save, 5, 10, 1, 2)

    self.btn_clear = QPushButton('清除訊息')
    self.btn_clear.clicked.connect(self.clear_log)
    self.main_layout.addWidget(self.btn_clear, 5, 9, 1, 1)

    self.btn_info = QPushButton('使用說明')
    self.btn_info.clicked.connect(self.show_helper)
    self.main_layout.addWidget(self.btn_info, 5, 0, 1, 2)

    self.chk_openfinished = QCheckBox("執行完畢後自動開啟")
    self.chk_openfinished.stateChanged.connect(self.on_auto_open)
    self.main_layout.addWidget(self.chk_openfinished, 5, 6, 1, 3)
    
    self.on_fast_select()

    self.auto_open = False
    self.active_buttons = [
      self.in_startdate, self.in_enddate, self.btn_fast_select, self.btn_save,
      self.cmb_fast_select
    ]

  def setup_helper_window(self):
    global AppVersion
    self.window_helper = QMessageBox()
    self.window_helper.setStyleSheet("QLabel{min-width: 600px; font-size: 14px}")
    self.window_helper.setWindowTitle("使用說明")
    self.window_helper.setStandardButtons(QMessageBox.Ok)
    self.window_helper.setText('''
    <p>程式版本: $TAG_SOFTWARE_VERSION$<p>
    <p>此程式可抓取iTHome資安週報的新聞以及子新聞內容</p>
    <hr>
    <b>UI 說明：</b>
    <table border=0 cellpadding=0 cellspacing=10>
    <tr><td>開始、結束日期：想要抓的新聞日期範圍</td></tr>
    <tr><td>快速選擇：依照選項快速設定開始與結束日期</td></tr>
    <tr><td>清除訊息：清除訊息欄位內的內容</td></tr>
    <tr><td>選擇儲存路徑：選擇欲輸出的 csv 路徑，如果檔案已存在且已被 Excel 開啟，需要將其關閉否則無法輸出；選擇確認後程式將會開始執行。</td></tr>
    </table>
    <hr>
    <b>※注意事項</b>
    <p>程式執行期間輸入欄位將會鎖定，不過仍可以選擇是否執行完成後自動開啟檔案。</p>
    <p>如果是使用 Excel 開啟輸出的 CSV 且 Excel 預設編碼不是 UTF-8；請使用匯入資料(文字檔)的方式開啟檔案並選擇 UTF-8 編碼。</p>
    '''.replace('$TAG_SOFTWARE_VERSION$', AppVersion))
    btn_close = self.window_helper.button(QMessageBox.Ok)
    btn_close.setText("關閉")
    self.window_helper.hide()

  def show_helper(self):
    self.window_helper.show()

  def _logging(self):
    global Mutex
    Mutex.lock()
    while not self.log_queue.empty():
      msg = self.log_queue.get()
      self.edit_log.append(msg)
    self.edit_log.moveCursor(QTextCursor.End)
    Mutex.unlock()

  def append_log(self, message):
    self.log_queue.put(message)
    self.request_logging.emit()

  def clear_log(self):
    self.edit_log.clear()

  def on_auto_open(self, signal):
    self.auto_open = True if signal > 0 else False

  def on_fast_select(self):
    current = datetime.now()
    qdate_ed = QDate(current.year, current.month, current.day)
    qdate_st = QDate(current.year, current.month, current.day)
    idx = self.cmb_fast_select.currentIndex()
    for i,key in enumerate(FastSelectOptions.keys()):
      if i != idx:
        continue
      val = FastSelectOptions[key]
      if val == 'ThisYear':
        qdate_st = QDate(current.year, 1, 1)
      elif val == 'LastYear':
        qdate_st = QDate(current.year-1, 1, 1)
        qdate_ed = QDate(current.year-1,12,31)
      elif val == 'ThisMonth':
        qdate_st = QDate(current.year, current.month, 1)
      else:
        tdelta  = timedelta(val)
        date_st = current - tdelta
        qdate_st = QDate(date_st.year, date_st.month, date_st.day)
      break
    self.in_startdate.setDate(qdate_st)
    self.in_enddate.setDate(qdate_ed)
    

  def execute_mainproc(self):
    self.append_log("程式開始運行...")
    save_path, _ = QFileDialog.getSaveFileName(self, '另存為...', './', '純文字檔 (*.txt);; CSV UTF-8 (逗號分隔) (*.csv);; 所有檔案 (*.*)')
    self.save_path = save_path
    if not save_path:
      return
    self.append_log(f"輸出路徑: {save_path}")
    if not is_file_writable(save_path):
      self.append_log("程式無法輸出至指定的檔案")
      QMessageBox.critical(None, 'Error', '無法輸出至指定的路徑, 請確認您有權限寫入且該檔案沒有被開啟!')
      return
    self.disable_buttons()
    self.thread = QThread()
    self.worker = QtCurlWorker(self, self.in_startdate.date(), self.in_enddate.date())
    self.worker.moveToThread(self.thread)
    self.thread.started.connect(self.worker.run)
    self.worker.finished.connect(self.on_worker_finished)
    self.worker.errored.connect(self.on_worker_errored)
    self.worker.aborted.connect(self.on_worker_aborted)
    self.thread.finished.connect(self.thread.deleteLater)
    self.thread.start()
    
  def on_worker_finished(self):
    print("Worker finished")
    self.append_log("程式執行完畢!")
    QMessageBox.information(self, 'Info', "程式執行完畢!", QMessageBox.Ok)
    self.terminate_worker()
    if self.auto_open:
      _ok = open_external_file(self.save_path)
      if not _ok:
        self.edit_log("自動開啟不支援當前版本的作業系統, 請手動開啟")
  
  def terminate_worker(self):
    self.thread.quit()
    self.worker.deleteLater()
    self.enable_buttons()
  
  def on_worker_aborted(self):
    print("Worker aborted")
    self.terminate_worker()
  
  def on_worker_errored(self):
    global LastError
    print("Worker errored")
    self.terminate_worker()
    if LastError:
      handle_exception(*LastError, window=MainWindow)

  def enable_buttons(self):
    for btn in self.active_buttons:
      btn.setDisabled(False)

  def disable_buttons(self):
    for btn in self.active_buttons:
      btn.setEnabled(False)

def start():
  global MainApp, MainWindow
  MainApp    = QApplication(sys.argv)
  MainWindow = MainGUI()
  MainWindow.show()
  sys.exit(MainApp.exec_())

sys._excepthook = sys.excepthook 
def exception_hook(exctype, value, traceback):
  print(exctype, value, traceback)
  sys._excepthook(exctype, value, traceback) 
  sys.exit(1) 
sys.excepthook = exception_hook 

if __name__ == '__main__':
  safe_execute_func(start)