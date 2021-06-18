from PyQt5.QtCore import Qt, QDateTime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QMessageBox, QFileDialog,
              QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QDateTimeEdit, 
              QTextEdit, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView)
import os
import sys
import traceback
import requests
from datetime import datetime
from bs4 import BeautifulSoup as BS
from threading import Thread

RootURI   = 'http://www.ithome.com.tw'
TargetURI = "http://www.ithome.com.tw/taxonomy/term/5971/all"
PageParam = 'page'

ContentClassPost = '.item'
ContentClassDate = '.post-at'
ContentClassTitle = '.title'
ContentClassTags  = '.category'
SigWorking  = 0
SigComplete = 1
SigError    = 2
SigAbort    = 3

MainWindow = None
MainApp    = None
Signal     = 0

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
    err_info = traceback.format_exc()
    handle_exception(err, err_info, MainWindow)
    print(f"An error occurred!\n{err_info}")
  return None

def handle_exception(err, errinfo, window=None):
  dmp_ok = dump_errorlog(err, errinfo)
  if window:
    QMessageBox.critical(None, 'Error', '運行過程中產生錯誤, 詳情請見資訊窗格')
    if dmp_ok:
      window.edit_log.append(' ')
      window.edit_log.append(f"運行過程中產生錯誤, 請將紀錄檔 `{dmp_ok}` 寄送給開發人員以排除錯誤")
      window.edit_log.append(f"錯誤內容: {err}")
      window.edit_log.append("同時請確認您的輸出路徑無誤")
    else:
      window.edit_log.append(f"運行過程中產生錯誤, 請將以下內容複製並寄送給開發人員以排除錯誤")
      window.edit_log.append(str(err))
      window.edit_log.append(errinfo)
      window.edit_log.append("同時請確認您的輸出路徑無誤")
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


### Main curl process
def is_internet_available():
  try:
    if requests.get('http://www.google.com', timeout=3):
      return True
  except Exception:
    pass
  return False

def get_page(page):
  return requests.get(f"{TargetURI}?{PageParam}={page}")

def parse_content(page):
  ret = []
  for post in list(page.select(ContentClassPost)):
    obj = {}
    # tags
    try:
      tags = post.select(ContentClassTags)[0].text
      _tr = str.maketrans({' ': ''})
      obj['tags'] = '|'.join([tag for tag in tags.translate(_tr).split('|') if len(tag) > 1])
    except Exception:
      pass
    # title & href
    try:
      node = post.select(ContentClassTitle)[0]
      obj['title'] = node.text
      obj['link']  = f"{RootURI}{node.select('a')[0]['href']}"
    except Exception:
      pass
    # date
    try:
      obj['date'] = post.select(ContentClassDate)[0].text.strip()
    except Exception:
      pass
    if obj:
      ret.append(obj)
  return ret

### GUI section
class MainGUI(QMainWindow):
  def __init__(self):
    super().__init__()
    self.resize(600, 400)
    self.setWindowTitle("iThome 資安周報爬蟲工具")
    self.setWindowOpacity(1)
    self.init_ui()
  
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
    self.main_layout.addWidget(self.cmb_fast_select, 0, 9, 1, 2)
    self.btn_fast_select = QPushButton('選擇')
    self.btn_fast_select.clicked.connect(self.on_fast_select)
    self.main_layout.addWidget(self.btn_fast_select, 0, 11, 1, 1)


    self.edit_log = QTextEdit(f'初始化成功; 目標網址: {TargetURI}')
    self.edit_log.setReadOnly(True)
    self.main_layout.addWidget(self.edit_log, 1, 0, 4, 12)
    
    self.btn_save = QPushButton('選擇儲存路徑')
    self.btn_save.clicked.connect(lambda: safe_execute_func(self.execute_mainproc))
    self.main_layout.addWidget(self.btn_save, 5, 10, 1, 2)

    self.chk_openfinished = QCheckBox("生成檔案後自動開啟")
    self.chk_openfinished.stateChanged.connect(self.on_auto_open)
    self.main_layout.addWidget(self.chk_openfinished, 5, 7, 1, 3)
    
    self.active_buttons = [
      self.in_startdate, self.in_enddate, self.btn_fast_select, self.btn_save,
      self.cmb_fast_select
    ]

  def on_auto_open(self, signal):
    self.auto_open = True if signal > 0 else False

  def on_fast_select(self):
    date = self.in_startdate.date()
    date = (date.year(), date.month(), date.day())

  def execute_mainproc(self):
    global Signal
    self.disable_buttons()
    self.edit_log.append("程式開始運行...")
    save_path, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'CSV UTF-8 (逗號分隔) (*.csv)')
    self.edit_log.append(f"輸出路徑: {save_path}")
    Signal = SigWorking
    while Signal == SigWorking:
      MainApp.processEvents()
      
    self.enable_buttons()

  def start_async(self):
    pass

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

if __name__ == '__main__':
  safe_execute_func(start)