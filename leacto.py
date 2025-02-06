import sys, os, base64
from PyQt6.QtWidgets import *
from PyQt6 import uic
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import pyqtSignal, pyqtSlot, QThread, QTimer #, QEvent, QSize
from threading import Lock
from core import LmsCore
from functools import wraps

NAME = 'Leacto'
VERSION = '1.0'
WINDOW_TITLE = f'{NAME} {VERSION}'
WINDOW_ICON = 'leacto.ico'

UI_MAINWINDOW = 'leacto.ui'
UI_SUBWINDOW = 'leacto_browser.ui'

BROWSER_REFRESH_RATE = 25 # Hz
BROWSER_REFRESH_DELAY = 1000 // BROWSER_REFRESH_RATE

BROWSER_WIDTH = 1080
BROWSER_HEIGHT = 640
BROWSER_SIZE = f'{BROWSER_WIDTH},{BROWSER_HEIGHT}'

# if PyInstaller bundled
_BUNDLED = getattr(sys, 'frozen', False)
if _BUNDLED:
    WINDOW_ICON = os.path.join(sys._MEIPASS, WINDOW_ICON)
    UI_MAINWINDOW = os.path.join(sys._MEIPASS, UI_MAINWINDOW)
    UI_SUBWINDOW = os.path.join(sys._MEIPASS, UI_SUBWINDOW)

# Load UI file
form_class = uic.loadUiType(UI_MAINWINDOW)[0]
form_class2 = uic.loadUiType(UI_SUBWINDOW)[0]

class Leacto(QMainWindow, form_class):
    course_signal = pyqtSignal(list)
    statusbar_signal = pyqtSignal(str)
    work_list = []

    class Worker(QThread):
        job_done = pyqtSignal(list)
        def __init__(self, func, connector = None, args = []):
            super().__init__()
            self.func = func
            self.connector = connector
            self.args = args
            if self.connector:
                self.job_done.connect(connector)
        def run(self):
            if self.connector:
                self.job_done.emit([self.func(*self.args)])
            else:
                self.func(*self.args)

    def breakEmission(func):
        @wraps(func)
        def wrapper(obj, return_list = []):
            func(obj, *return_list)
        return wrapper

    def work(self, func, connector = None, args = [], start_msg = '', end_msg = ''):
        if not self.worker or not self.worker.isRunning():
            def _work(*args):
                self.set_statusbar(start_msg)
                result = func(*args)
                self.set_statusbar(end_msg)
                return result
            self.worker = Leacto.Worker(_work, connector, args)
            self.worker.start()

    def __init__(self):
        super().__init__()
        self.lock = Lock()
        self.core = None
        self.worker = None
        self.build_ui()
        self.show()
        self.statusbar_signal.connect(self.on_set_statusbar)
        self.work(self.init_core, self.on_init_core, start_msg = '브라우저 준비 중...')
        self.course_signal.connect(self.on_course)


    def __del__(self):
        self.close()

    def init_core(self):
        with self.lock:
            self.core = LmsCore(size=BROWSER_SIZE)
        return True

    @pyqtSlot(list)
    @breakEmission
    def on_init_core(self, success):
        if success:
            self.btnLogin.setEnabled(True)
            self.chkBrowser.setEnabled(True)

    def click_login(self):
        self.work(self.login, self.on_login, start_msg = '로그인 중...')

    def login(self):
        id = self.lineId.text()
        pw = self.linePw.text()
        if id and pw:
            with self.lock:
                return self.core.login(self.lineUrl.text(), id, pw)

    @pyqtSlot(list)
    @breakEmission
    def on_login(self, success):
        print('on_login', success)
        if success:
            self.tabLogin.setEnabled(False)
            self.tabWidget.setCurrentWidget(self.tabCourse)
            self.work(self.get_courselist, self.on_get_courselist)
        else:
            self.set_statusbar('로그인 실패')

    def get_courselist(self):
        try:
            return self.core.get_course()
        except:
            self.tabWidget.setCurrentWidget(self.tabLogin)

    @pyqtSlot(list)
    @breakEmission
    def on_get_courselist(self, course_list):
        self.lstCourse.clear()
        self.course_info = course_list
        for course in self.course_info:
            self.lstCourse.addItem(course['text'])
        self.tabCourse.setEnabled(True)

    def doubleclick_course(self):
        self.tabCourse.setEnabled(False)
        self.btnCloseCourse.setEnabled(True)
        self.tabWidget.setCurrentWidget(self.tabProgress)
        self.work(self.enter_course, self.on_finish_course)
    
    def enter_course(self):
        # self.course_signal.connect(self.on_course)    # __init__
        idx = self.lstCourse.currentRow()
        return self.core.enter_course(idx, self.course_signal)

    @pyqtSlot(list)
    @breakEmission
    def on_finish_course(self, success = True):
        try:
            self.on_login([True])
        except:
            # Session Expired
            self.tabLogin.setEnabled(True)
            self.tabWidget.setCurrentWidget(self.tabLogin)
            self.set_statusbar('세션 종료 - 로그인 필요')

    @pyqtSlot(list)
    def on_course(self, emission):
        match(emission[0]):
            case 1:
                self.lblCourseInfo1.setText(emission[1])
            case 2:
                self.lblCourseInfo2.setText(emission[1])
            case 0:
                progress, played, length = emission[1:]
                self.pbProgress.setValue(int(progress))
                if played != '':
                    self.lblCourseInfo3.setText(f'{played}/{length}')
                else:
                    self.lblCourseInfo3.setText('')
            case -1:    # Stopped
                self.clear_course()
                self.on_finish_course()

    def clear_course(self):
        self.on_course([1, ''])
        self.on_course([2, ''])
        self.on_course([0, 0, '', ''])

    def build_ui(self):
        self.setupUi(self)
        self.setWindowTitle(WINDOW_TITLE)
        self.setWindowIcon(QIcon(WINDOW_ICON))
        # self.statusbar.setSizeGripEnabled(False)
        self.browser = LeactoBrowserWin(self)
        self.chkBrowser.setEnabled(False)
        self.chkBrowser.clicked.connect(self.toggle_browser)
        self.linePw.returnPressed.connect(self.click_login)
        self.btnLogin.clicked.connect(self.click_login)
        self.lstCourse.doubleClicked.connect(self.doubleclick_course)
        self.btnCloseCourse.setEnabled(False)
        self.btnCloseCourse.clicked.connect(self.stop_course)

    def toggle_browser(self, checked):
        self.browser.show() if checked else self.browser.hide()

    def stop_course(self):
        if self.core:
            self.core.stop = True
            self.btnCloseCourse.setEnabled(False)

    def set_statusbar(self, msg = ''):
        self.statusbar_signal.emit(msg)

    @pyqtSlot(str)
    def on_set_statusbar(self, msg):
        self.lblStatusbar.setText(msg)

    def grab_screen(self):
        return base64.b64decode(self.core.driver.get_screenshot_as_base64())

    def closeEvent(self, _):
        self.hide()
        self.browser.hide()
        if self.worker and self.worker.isRunning():
            self.worker.wait()
        self.work(self.close)

    def close(self):
        if self.core:
            self.core.close()
            self.core = None
            self.browser = None

class LeactoBrowserWin(QMainWindow, form_class2):
    def __init__(self, main):
        super().__init__()
        self.build_ui()
        self.main = main
        self.screen = QPixmap()
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_screen)
        self.timer.setInterval(BROWSER_REFRESH_DELAY)
        self.installEventFilter(self)

    def build_ui(self):
        self.setupUi(self)
        self.setWindowTitle(f'{WINDOW_TITLE} - Browser')
        self.setWindowIcon(QIcon(WINDOW_ICON))

    def refresh_screen(self):
        try:
            self.screen.loadFromData(self.main.grab_screen())
            self.label.setPixmap(self.screen)
        except:
            pass

    def showEvent(self, _):
        self.timer.start()
    def hideEvent(self, _):
        self.timer.stop()
    def closeEvent(self, _):
        self.main.chkBrowser.setChecked(False)

    def wheelEvent(self, event):
        self.main.core.scroll(0, event.angleDelta().y() // 3)

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    window = Leacto()

    if _BUNDLED:
        sys.exit(app.exec())
    else:
        window.setGeometry(1520,800,0,0)
        self = window

