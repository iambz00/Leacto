from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.common.exceptions as SeleniumException
import time, os, urllib3, sys

_BUNDLED = getattr(sys, 'frozen', False)

def querySelector(root, query):
    return root.find_element(By.CSS_SELECTOR, query)
def querySelectorAll(root, query):
    return root.find_elements(By.CSS_SELECTOR, query)

class LmsCore:
    def __init__(self, headless = True, size = "1080,580", mute = True):
        self.delay = 0.3
        self.courseList = []
        self.stop = False
        self.mute = True
        os.environ["WDM_SSL_VERIFY"] = "0"
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        options = webdriver.ChromeOptions()
        options.add_argument("--log-level=3")
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--window-size=" + size)
        if mute:
            options.add_argument("--mute-audio")
        options.add_argument("--disable-gpu")
        options.add_argument("--hide-scrollbars")
        options.add_argument("--autoplay-policy=no-user-gesture-required")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
        cdm = ChromeDriverManager()
        driver_path = os.path.join(os.path.dirname(cdm.install()), "chromedriver.exe")
        major_version = cdm.driver.get_driver_version_to_download().split('.')[0]

        user_agent = f'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{major_version}.0.0.0 Safari/537.36'
        print(f"* Set User-Agent: {user_agent}")
        options.add_argument("user-agent=" + user_agent)
        self.driver = webdriver.Chrome(service=ChromeService(driver_path, popen_kw={"creation_flags":134217728}), options=options)
        # popen_kw is needed to suppress 'DevTools listening on' Message
        self.driver.implicitly_wait(1)

        self.hMain = self.driver.current_window_handle
        self.wait = WebDriverWait(self.driver, timeout=10)

    def __del__(self):
        if self.driver:
            self.close()

    def set_base_url(self, url):
        self.url = url

    def close_popups(self):
        while len(self.driver.window_handles) > 1:
            for handle in self.driver.window_handles:
                if handle != self.hMain:
                    self.driver.switch_to.window(handle)
                    self.driver.close()
        self.driver.switch_to.window(self.hMain)

    def login(self, url='', userID=None, userPW=None):
        self.set_base_url(url)
        try:
            print("* 로그인 중...")
            self.driver.get(self.url + "/system/login/login.do")
            self.close_popups()
            login_id = querySelector(self.driver, "#userInputId")
            login_id.click()
            webdriver.ActionChains(self.driver).pause(0.5)\
                .send_keys(userID).send_keys(Keys.TAB).pause(0.5)\
                .send_keys(userPW).pause(0.5)\
                .send_keys(Keys.ENTER).pause(1)\
                .perform()
        except Exception as e:
            print("* 실패")
            return False
        else:
            self.close_popups()
            try:
                if self.driver.current_url.find("login") > 0:
                    raise Exception()
                print("* 성공")
                return True
            except SeleniumException.UnexpectedAlertPresentException as e:
                print("* 실패: " + e.alert_text)
                try:
                    self.driver.switch_to.alert.accept()
                except:
                    pass
                return False
            except Exception as e:
                print("* 실패")
                return False

    def get_course(self):
        try:
            print("\n* 과정 선택")
            self.driver.get(self.url + "/lh/ms/cs/atnlcListView.do?menuId=3000000101")

            self.driver.find_elements(By.CSS_SELECTOR, "#crseList > li")
            courses = self.driver.find_elements(By.CSS_SELECTOR, "#crseList > li")

            self.courseList = []
            for i in range(len(courses)):
                text = courses[i].find_element(By.CSS_SELECTOR, "a.title").text
                if len(text) > 40:
                    text = text[:36] + "..."
                button = ""
                progress = int(float(courses[i].find_element(By.CSS_SELECTOR, "ul.progress_list_wrap p.num").text.replace('%', '')))
                for abtn in courses[i].find_elements(By.CSS_SELECTOR, "a"):
                    if abtn.text == "이어보기" or abtn.text == "학습하기":
                        button = abtn
                        break
                self.courseList.append({ 'text': f"  [{i+1}] {text}", 'obj': button, 'progress': progress})
            for i in range(len(self.courseList)):
                print(self.courseList[i]['text'])
            return self.courseList
        except SeleniumException.UnexpectedAlertPresentException as e:
            self.driver.switch_to.alert.accept()
            return None

    def enter_course(self, num, signal):
        if num not in range(len(self.courseList)):
            num = 0
        num_windows = len(self.driver.window_handles)
        try:
            self.courseList[num]['obj'].click()
            time.sleep(2)
        except SeleniumException.UnexpectedAlertPresentException as e:
            print("\n* 경고 - " + e.alert_text)
            self.driver.switch_to.alert().accept()
            e.add_note(e.alert_text)
            raise
        except Exception as e:
            print("\n* 오류 발생 - " + str(e))
            raise
        else:
            self.hLearn = self.get_new_window(num_windows)
            self.driver.switch_to.window(self.hLearn)
            return self.learn(signal)

    def get_new_window(self, num_windows_before):
        self.wait.until(EC.number_of_windows_to_be(num_windows_before + 1))
        return self.driver.window_handles[-1]

    def return_to_main(self):
        self.driver.switch_to.window(self.hMain)

    def learn(self, signal):
        print("\n* 수강 시작")
        current_subject = " "
        current_section = " "
        current_subsect = " "
        current_progress = 0.0
        # Get First Button
        startbutton = querySelector(self.driver, 'a.btn_learning_list')
        if startbutton.is_displayed():
            startbutton.click()
            time.sleep(2)
        try:
            while not self.stop:
                vjs_controlbar = querySelector(self.driver, "div.vjs-control-bar")
                if str(vjs_controlbar.value_of_css_property("opacity")) != "1":
                    # Show video player control bar
                    # Set best playback rate. The rate is saved in 'Local storage'.
                    self.driver.execute_script('''
                        document.querySelector('div.vjs-control-bar').style.opacity = 1
                        document.querySelector('.vjs-playback-rate.vjs-menu-button .vjs-menu-item').click()
                    ''')
                time.sleep(self.delay)
                subject = querySelector(self.driver, "div.class_list p.title_box").text.strip()
                section = querySelector(self.driver, "div.class_list_box.ing li.play div a").text.strip()
                if not section or section == "학습하기":
                    section = querySelector(self.driver, "div.class_list_box.ing p").text.strip()
                subsect = querySelector(self.driver, "#page-info").text.strip()

                progress = float(querySelector(self.driver, "#lx-player div.vjs-progress-holder").get_attribute("aria-valuenow"))
                progress_time = querySelector(self.driver, "#lx-player div.vjs-progress-holder").get_attribute("aria-valuetext").strip().split()
                if len(progress_time) > 0:
                    played = progress_time[0]
                    length = progress_time[2]

                if subject[1:3] == '차시':
                    subject = subject[:3] + '> ' + subject[3:]
                elif subject[2:4] == '차시':
                    subject = subject[:4] + '> ' + subject[4:]
                if subject and subject != current_subject:
                    print(f"\r[차시]: {subject}")
                    signal.emit([1, f"{subject}"])
                    current_subject = subject
                if section and section != current_section:
                    print(f"\r  [강의]: {section} [{subsect}] {length}")
                    signal.emit([2, f"{section} [{subsect}]"])
                    current_section = section
                    current_subsect = subsect
                if subsect and subsect != current_subsect:
                    print(f"\r  [강의]: {section} [{subsect}] {length}")
                    signal.emit([2, f"{section} [{subsect}]"])
                    current_subsect = subsect
                is_quizpage = querySelector(self.driver, "#quizPage").is_displayed()
                playbutton = querySelector(self.driver, "button.vjs-big-play-button")
                if progress == current_progress:
                    if is_quizpage:
                        print(f"\r  [퀴즈]")
                        signal.emit([2, "- 퀴즈 -"])
                        current_progress = 100.0
                    elif played == length:
                        current_progress = 100.0
                        time.sleep(1)
                    elif playbutton.is_displayed():
                        print(f"\r! 재생 시작")
                        playbutton.click()
                else:
                    current_progress = progress
                signal.emit([0, current_progress, played, length])
                if current_progress >= 100.0:
                    try:
                        popup = querySelector(self.driver, 'div.popup_wrap p.desc')
                        if popup.is_displayed():
                            print(f'* {popup.text}')
                            self.stop = True
                            break
                    except: pass
                    self.driver.execute_script("next_ScoBtn()")
                    time.sleep(1)

        # except SeleniumException.UnexpectedAlertPresentException as e:
        #     print("\n* 경고 - " + e.alert_text)
        #     return_status = e.alert_text
        # except SeleniumException.JavascriptException as e:
        #     print("\n* 수강 완료")
        #     return_status = 'FINISH'
        except Exception as e:
            alert_text = ''
            try:
                alert_text = self.driver.switch_to.alert.text
                self.driver.switch_to.alert.accept()
            except:
                pass
            print(f"\n* 오류: {alert_text}")
            # Don't return here

        # Stop Button Pushed
        signal.emit([-1])
        if self.driver.service.is_connectable():
            self.stop = False
            self.driver.close()
            self.return_to_main()
        return True

    def close(self):
        print("\n* 종료합니다.")
        if self.driver.service.is_connectable():
            self.driver.quit()
            self.driver = None

    def scroll(self, h, v):
        webdriver.ActionChains(self.driver).scroll_by_amount(h, -v).perform()
