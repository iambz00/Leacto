import sys, os, base64
from PyQt6.QtWidgets import *
from PyQt6 import uic
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import pyqtSignal, pyqtSlot, QThread, QTimer #, QEvent, QSize
from threading import Lock
from core import LmsCore
from functools import wraps

NAME = 'Leacto'
VERSION = '1.1'
WINDOW_TITLE = f'{NAME} {VERSION}'
WINDOW_ICON = 'leacto.ico'

UI_MAINWINDOW = 'leacto.ui'
UI_SUBWINDOW = 'leacto_browser.ui'

BROWSER_REFRESH_RATE = 25 # Hz
BROWSER_REFRESH_DELAY = 1000 // BROWSER_REFRESH_RATE

BROWSER_WIDTH = 1280
BROWSER_HEIGHT = 920
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
        def _work(*args):
            self.set_statusbar(start_msg)
            result = func(*args)
            self.set_statusbar(end_msg)
            return result
        worker = Leacto.Worker(_work, connector, args)
        self.workers = [w for w in self.workers if not w.isFinished()]
        self.workers.append(worker)
        worker.start()


    def __init__(self):
        super().__init__()
        self.lock = Lock()
        self.core = None
        self.worker = None
        self.workers = []
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

    def login(self):
        def _login():
            id = self.lineId.text()
            pw = self.linePw.text()
            if id and pw:
                with self.lock:
                    return self.core.login(self.lineUrl.text(), id, pw)
        self.work(_login, self.on_login, start_msg = '로그인 중...')


    @pyqtSlot(list)
    @breakEmission
    def on_login(self, success):
        if success:
            self.tabLogin.setEnabled(False)
            self.tabWidget.setCurrentWidget(self.tabCourse)
            def get_courselist():
                try:
                    return self.core.get_course()
                except:
                    self.tabWidget.setCurrentWidget(self.tabLogin)
                    return []
            self.work(get_courselist, self.on_get_courselist)
        else:
            self.set_statusbar('로그인 실패')

    @pyqtSlot(list)
    @breakEmission
    def on_get_courselist(self, course_list):
        self.lstCourse.clear()
        self.course_info = course_list
        for course in self.course_info:
            self.lstCourse.addItem(f'{course['text']}')
        self.tabCourse.setEnabled(True)

    def select_course(self):
        selected = self.lstCourse.currentRow()
        if self.course_info[selected]['progress'] < 100:
            self.btnStart.setEnabled(True)
            self.btnHome.setEnabled(False)
        else:
            self.btnStart.setEnabled(False)
            self.btnHome.setEnabled(True)
        self.btnSurvey.setEnabled(False)
        self.btnExam.setEnabled(False)

    def start_course(self):
        self.tabCourse.setEnabled(False)
        self.btnCloseCourse.setEnabled(True)
        self.tabWidget.setCurrentWidget(self.tabProgress)
        def _start_course():
            idx = self.lstCourse.currentRow()
            return self.core.enter_course(idx, self.course_signal)
        self.work(_start_course, self.on_finish_course)

    def open_home(self):
        self.work(self.core.go_course_home, args=[self.lstCourse.currentRow()], connector=self.on_open_home)

    def on_open_home(self):
        self.btnHome.setEnabled(False)
        self.btnSurvey.setEnabled(True)
        self.btnExam.setEnabled(True)

    def open_survey(self):
        self.work(self.core.go_survey, connector=self.on_open_survey)

    @pyqtSlot(list)
    @breakEmission
    def on_open_survey(self, _):
        self.browser.survey_mode = True
        self.browser.exam_mode = False
        self.browser.show()

    def open_exam(self):
        self.work(self.core.go_exam, connector=self.on_open_exam)

    @pyqtSlot(list)
    @breakEmission
    def on_open_exam(self, _):
        pass

    def return_to_list(self):
        def _clear_windows():
            self.core.clear_windows()
            return True
        self.work(_clear_windows, connector=self.on_login)

    @pyqtSlot(list)
    @breakEmission
    def on_finish_course(self, _):
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
        self.linePw.returnPressed.connect(self.login)
        self.btnLogin.clicked.connect(self.login)

        # self.lstCourse.doubleClicked.connect(self.doubleclick_course)
        self.lstCourse.currentItemChanged.connect(self.select_course)
        self.btnStart .clicked.connect(self.start_course)
        self.btnHome  .clicked.connect(self.open_home)
        self.btnSurvey.clicked.connect(self.open_survey)
        self.btnExam  .clicked.connect(self.open_exam)
        self.btnList  .clicked.connect(self.return_to_list)

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
        self.setFixedSize(1280, 780)
        self.screen = QPixmap()
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_screen)
        self.timer.setInterval(BROWSER_REFRESH_DELAY)
        self.installEventFilter(self)
        self.survey_mode = False
        self.exam_mode = False

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

    def check(self, num):
        if self.survey_mode:
            return self.main.core.survey_check(num)
        elif self.exam_mode:
            return self.main.core.exam_check(num)
        return False

    def move(self, forward = True):
        if self.survey_mode:
            return self.main.core.survey_move(forward)
        elif self.exam_mode:
            return self.main.core.exam_move(forward)
        return False

    def showEvent(self, _):
        self.main.chkBrowser.setChecked(True)
        self.timer.start()
    def hideEvent(self, _):
        self.main.chkBrowser.setChecked(False)
        self.timer.stop()
    def closeEvent(self, _):
        self.main.chkBrowser.setChecked(False)

    def keyPressEvent(self, event):
        if self.survey_mode or self.exam_mode:
            key = event.key()
            if key == 49: # 1
                self.check(1)
            elif key == 50: # 2
                self.check(2)
            elif key == 51: # 3
                self.check(3)
            elif key == 52: # 4
                self.check(4)
            elif key == 53: # 5
                self.check(5)
            elif key == 16777234: # Left
                self.move(False)
            elif key == 16777236: # Right
                self.move()
            elif key == 16777235: # Up
                self.main.core.scroll(0, 100)
            elif key == 16777237: # Down
                self.main.core.scroll(0, -100)
            
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

