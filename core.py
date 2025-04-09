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
def log(*args, **kwargs):
    print(*args, **kwargs)
def querySelector(root, query):
    return root.find_element(By.CSS_SELECTOR, query)
def querySelectorAll(root, query):
    return root.find_elements(By.CSS_SELECTOR, query)

class LmsCore:
    def __init__(self, headless = True, size = "1080,580", mute = True):
        self.delay = 0.3
        self.courseList = []
        self.stop = False
        self.pause = False
        self.mute = True
        os.environ["WDM_SSL_VERIFY"] = "0"
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("headless=new")
        if mute:
            options.add_argument("mute-audio")
        options.add_argument("--log-level=3")
        options.add_argument("window-size=" + size)
        options.add_argument("disable-gpu")
        options.add_argument("hide-scrollbars")
        options.add_argument("autoplay-policy=no-user-gesture-required")
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.password_manager_leak_detection": False,   # Suppress 'Password Chane' warning
        }
        options.add_experimental_option("prefs", prefs)
    
        cdm = ChromeDriverManager()
        driver_path = os.path.join(os.path.dirname(cdm.install()), "chromedriver.exe")
        major_version = cdm.driver.get_driver_version_to_download().split('.')[0]
        user_agent = f'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{major_version}.0.0.0 Safari/537.36'
        # print(f"* Set User-Agent: {user_agent}")
        options.add_argument("--user-agent=" + user_agent)

        self.driver = webdriver.Chrome(service=ChromeService(driver_path, popen_kw={"creation_flags":134217728}), options=options)
        # popen_kw is needed to suppress 'DevTools listening on' Message
        self.driver.implicitly_wait(1)

        self.window_handle_main = self.driver.current_window_handle
        self.wait = WebDriverWait(self.driver, timeout=10)

    def __del__(self):
        if self.driver:
            self.close()

    def set_base_url(self, url):
        self.url = url

    def clear_windows(self):
        # Close all windows except main window
        while len(self.driver.window_handles) > 1:
            for handle in self.driver.window_handles:
                if handle != self.window_handle_main:
                    self.driver.switch_to.window(handle)
                    self.driver.close()
        self.driver.switch_to.window(self.window_handle_main)

    def login(self, url='', userID=None, userPW=None):
        self.set_base_url(url)
        try:
            log("* 로그인")
            self.driver.get(self.url + "/system/login/login.do")
            self.clear_windows()
            login_id = querySelector(self.driver, "#userInputId")
            login_id.click()
            webdriver.ActionChains(self.driver).pause(0.2)\
                .send_keys(userID).send_keys(Keys.TAB).pause(0.2)\
                .send_keys(userPW).pause(0.2)\
                .send_keys(Keys.ENTER).pause(0.2)\
                .perform()
        except Exception as e:
            log("* 실패 - " + str(e))
            return False
        else:
            self.clear_windows()
            try:
                if self.driver.current_url.find("login") > 0:
                    raise Exception()
                log("* 성공")
                return True
            except SeleniumException.UnexpectedAlertPresentException as e:
                log("* 실패: " + e.alert_text)
                try:    self.driver.switch_to.alert.accept()
                except: pass
                return False
            except Exception as e:
                log("* 실패")
                return False

    def get_course(self):
        self.courseList = []
        try:
            log("\n* 과정 선택")
            self.driver.get(self.url + "/lh/ms/cs/atnlcListView.do?menuId=3000000101")

            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#crseList > li")))
            self.driver.implicitly_wait(1)
            courses = querySelectorAll(self.driver, "#crseList > li")

            self.courseList = []
            for i in range(len(courses)):
                title = querySelector(courses[i],"a.title").text
                # if len(title) > 40:
                #     title = title[:36] + "..."
                obj_start = None
                obj_home  = None
                progress = int(float(querySelector(courses[i],"ul.progress_list_wrap p.num").text.replace('%', '')))
                for abtn in querySelectorAll(courses[i], "a"):
                    if abtn.text == "이어보기" or abtn.text == "학습하기":
                        obj_start = abtn
                    if abtn.text == "강의실 홈":
                        obj_home = abtn
                self.courseList.append({
                    'text': f"  [{i+1}] [{progress}%] {title}",
                    'obj_start': obj_start,
                    'obj_home': obj_home,
                    'progress': progress,
                })
            for i in range(len(self.courseList)):
                print(self.courseList[i]['text'])
        except SeleniumException.UnexpectedAlertPresentException as e:
            log("* 실패: " + e.alert_text)
            try:    self.driver.switch_to.alert.accept()
            except: pass
            self.courseList = []
        finally:
            return self.courseList

    def enter_course(self, num, signal):
        num_windows = len(self.driver.window_handles)
        try:
            self.courseList[num]['obj_start'].click()
        except SeleniumException.UnexpectedAlertPresentException as e:
            log("* 실패: " + e.alert_text)
            try:    self.driver.switch_to.alert.accept()
            except: pass
            raise
        except Exception as e:
            log("\n* 오류 발생 - " + str(e))
            raise
        else:
            self.driver.switch_to.window(self.get_new_window(num_windows))
            return self.learn(signal)

    def get_new_window(self, num_windows_before):
        self.wait.until(EC.number_of_windows_to_be(num_windows_before + 1))
        return self.driver.window_handles[-1]
    
    def return_to_main(self):
        self.driver.switch_to.window(self.window_handle_main)

    def learn(self, signal):
        log("\n* 수강 시작")
        current_subject = " "
        current_section = " "
        current_subsect = " "
        current_progress = 0.0
        # Get First Button
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.btn_learning_list')))
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
                        if not self.pause:
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
        #     log("\n* 경고 - " + e.alert_text)
        #     return_status = e.alert_text
        # except SeleniumException.JavascriptException as e:
        #     log("\n* 수강 완료")
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
        log("\n* 종료합니다.")
        if self.driver.service.is_connectable():
            self.driver.quit()
            self.driver = None

    def scroll(self, h, v):
        webdriver.ActionChains(self.driver).scroll_by_amount(h, -v).perform()

    def go_course_home(self, num):
        log("\n* 강의실 홈")
        num_windows = len(self.driver.window_handles)
        try:
            self.courseList[num]['obj_home'].click()
            time.sleep(2)
        except SeleniumException.UnexpectedAlertPresentException as e:
            log("* 실패: " + e.alert_text)
            try:    self.driver.switch_to.alert.accept()
            except: pass
            return False
        except Exception as e:
            log("\n* 오류 발생 - " + str(e))
            return False
        else:
            self.driver.switch_to.window(self.get_new_window(num_windows))
            return True

    def go_survey(self):
        log("** 설문조사")
        querySelector(self.driver, 'a#surveyBtnControl').click()
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.survey button.start')))
        self.driver.implicitly_wait(1)
        log("*** 시작")
        querySelector(self.driver, 'div.survey button.start').click()
        return True

    def survey_check(self, num):
        questions = querySelectorAll(querySelector(self.driver, '#qusnInfo'), 'label')
        if num <= len(questions):
            log(f'  [{num}] 체크')
            questions[num-1].click()
            return True
        return False

    def survey_move(self, forward = True):
        if forward:
            try:
                querySelector(self.driver, 'button#btn_next').click()
            except:
                # if last, submit
                pass
        else:
            try:
                querySelector(self.driver, 'button#btn_before').click()
            except:
                # if first, no before
                pass
        return True

    def go_exam(self):
        log("** 시험")
        querySelector(self.driver, '#examPageLi a').click()
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.exam button.start')))
        self.driver.implicitly_wait(1)
        log("*** 시작")
        querySelector(self.driver, 'div.exam button.start').click()
        return True

    def exam_check(self, num):
        questions = querySelectorAll(querySelector(self.driver, 'div.exam_con'), 'label')
        if num <= len(questions):
            questions[num-1].click()
            return True
        return False

    def exam_move(self, forward = True):
        if forward:
            querySelectorAll(self.driver, 'button.btn_basic_small')[-1].click()
        else:
            querySelectorAll(self.driver, 'button.btn_basic_small')[0].click()
        return True

    def get_answer_sheet(self):
        def indexOf(lst, e):
            try:
                return lst.index(e)
            except:
                return -1
        answer_sheet = []
        try:
            while True:
                answers = querySelectorAll(self.driver, '#layer .exam_con ol > li')
                for i in range(len(answers)):
                    if indexOf(answers[i].get_attribute('class').split(), 'correct') > -1:
                        answer_sheet.append(i+1)
                        break
                control_buttons = querySelectorAll(self.driver, '#layer a.btn_basic')
                if control_buttons[0].text.find('다음') > -1:
                    control_buttons[0].click()
                else:
                    control_buttons[1].click()
                webdriver.ActionChains(self.driver).pause(1).perform()
        except:
            pass
        return answer_sheet
