# token 자동삭제ver
# https://docs.google.com/spreadsheets/d/1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos/edit?gid=1906694512#gid=1906694512

import os
import sys

# auth.py 모듈 경로 추가
sys.path.append(os.path.expanduser("~/Documents/github_cloud/module_auth"))

import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import time
import random
import pyperclip
import pyautogui
from auth import get_credentials
from googleapiclient.discovery import build
import atexit
from naver_message_module import InstagramMessageTemplate

# 로깅 설정 - INFO 레벨 이상만 표시
logging.basicConfig(
    level=logging.INFO,
    format='%(message)s'  # 메시지만 간단히 표시
)

SPREADSHEET_ID = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos"

def get_data_from_sheets():
    logging.info("\n=== 메일 발송 준비를 시작합니다 ===")
    try:
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)

        # DB 목록 먼저 가져오기
        db_list = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='아이디보드!F2:F'  # F열의 DB 목록
        ).execute().get('values', [])
        
        # 유효한 DB 목록 생성 (빈 셀 제외)
        valid_dbs = [db[0] for db in db_list if db]

        # 스프레드시트의 모든 시트 정보 가져오기
        sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', '')
        
        # 메일 템플릿이 있는 시트 목록 출력
        print("\n=== 사용할 메일 템플릿 시트를 선택하세요 ===")
        template_sheets = []
        for i, sheet in enumerate(sheets, 1):
            sheet_title = sheet.get('properties', {}).get('title', '')
            # '_'로 시작하는 시트, '아이디보드' 시트, '완료'가 포함된 시트, '사용법' 시트, DB 목록에 있는 시트는 제외
            if (not sheet_title.startswith('_') 
                and sheet_title != '아이디보드'
                and sheet_title != '사용법'
                and '완료' not in sheet_title
                and sheet_title not in valid_dbs):
                print(f"{len(template_sheets) + 1}. {sheet_title}")
                template_sheets.append(sheet_title)

        # 사용자로부터 시트 선택 받기
        while True:
            try:
                selected_num = int(input("\n번호를 입력하세요: ")) - 1
                if 0 <= selected_num < len(template_sheets):
                    selected_sheet = template_sheets[selected_num]
                    break
                else:
                    print("올바른 번호를 입력해주세요.")
            except ValueError:
                print("숫자를 입력해주세요.")

        # DB 목록 출력 및 선택
        print("\n=== 사용할 DB를 선택하세요 ===")
        for i, db_name in enumerate(valid_dbs, 1):
            print(f"{i}. {db_name}")

        # 사용자로부터 DB 선택 받기
        while True:
            try:
                selected_db_num = int(input("\n번호를 입력하세요: ")) - 1
                if 0 <= selected_db_num < len(valid_dbs):
                    selected_db = valid_dbs[selected_db_num]
                    break
                else:
                    print("올바른 번호를 입력해주세요.")
            except ValueError:
                print("숫자를 입력해주세요.")

        # 아이디보드에서 계정 정보 가져오기
        accounts = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='아이디보드!A1:D'
        ).execute().get('values', [])

        # 헤더를 제외한 계정 목록 출력
        print("\n=== 사용할 계정 번호를 선택하세요 ===")
        for account in accounts[1:]:  # 첫 번째 행(헤더) 제외
            note = f" ({account[3]})" if len(account) > 3 and account[3] else ""  # D열 값이 있으면 괄호와 함께 표시
            print(f"{account[0]}. {account[1]}{note}")
        
        # 사용자 입력 받기
        while True:
            try:
                selected_num = input("\n번호를 입력하세요: ")
                # 선택된 계정 찾기
                selected_account = None
                for account in accounts[1:]:
                    if account[0] == selected_num:
                        selected_account = account
                        break
                
                if selected_account:
                    user_id = selected_account[1]
                    user_pw = selected_account[2]
                    break
                else:
                    print("올바른 번호를 입력해주세요.")
            except Exception as e:
                print("올바른 번호를 입력해주세요.")

        # 제목과 본문 템플릿 따로 가져오기
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        
        # 제목 템플릿 가져오기 (1행만)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{selected_sheet}!B1:D1'
        ).execute()
        email_titles = result.get('values', [[]])[0]
        
        # 본문 템플릿은 기존 방식대로
        message_template = InstagramMessageTemplate(SPREADSHEET_ID, selected_sheet)
        email_contents = message_template.get_message_templates()

        # 수신자 데이터 가져오기 (B열부터 D열까지)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{selected_db}!B2:D'
        ).execute()
        
        all_data = result.get('values', [])
        recipient_data = []
        skipped_count = 0

        # D열 데이터가 없는 행만 필터링
        for row_idx, row in enumerate(all_data, 2):  # 2부터 시작 (헤더 제외)
            if len(row) < 3 or not row[2].strip():  # D열이 비어있는 경우만 포함
                if len(row) >= 2:  # B, C열 데이터가 있는 경우
                    recipient_data.append((row[0], row[1], row_idx))  # 행 번호도 저장
                    logging.info(f"발송 대상: {row[1]}")
            else:
                skipped_count += 1
                logging.info(f"이미 발송됨 (건너뜀): {row[1]}")

        if skipped_count > 0:
            logging.info(f"\n이미 발송된 {skipped_count}개의 메일을 제외했습니다.")
        logging.info(f"총 {len(recipient_data)}개의 메일을 발송할 예정입니다.\n")

        logging.info("\n=== 메일 발송 준비가 완료되었습니다 ===")
        return email_titles, email_contents, user_id, user_pw, recipient_data, selected_db
    except Exception as e:
        logging.error(f"\n=== 오류가 발생했습니다: {str(e)} ===")
        raise

def create_driver():
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        
        # 불필요한 로그 숨기기
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_argument("--log-level=3")  # fatal 로그만 표시
        chrome_options.add_argument("--silent")
        
        # 자동화 표시 숨기기
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # WebGL 관련 경고 제거
        chrome_options.add_argument("--disable-gpu")  # GPU 하드웨어 가속 비활성화
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument('--disable-gpu-sandbox')
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--disable-webgl")
        chrome_options.add_argument("--disable-webgl2")
        
        # DevTools 메시지 숨기기
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument("--disable-in-process-stack-traces")
        
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        logging.error(f"\n=== 브라우저 실행 중 오류가 발생했습니다: {str(e)} ===")
        raise

def prevent_browser_close(driver):
    def keep_browser_open():
        try:
            if driver.service.process:
                driver.execute_script("Object.defineProperty(window, 'onbeforeunload', { value: function() { return true; } });")
        except Exception:
            pass
    return keep_browser_open

def update_sheet_status(service, spreadsheet_id, sheet_name, row_index, sent_time, status):
    """
    스프레드시트의 특정 행의 D열과 E열에 발송 시각과 상태를 업데이트합니다.
    row_index는 2부터 시작 (헤더 제외)
    """
    try:
        # '읽지않음' 상태를 '발송완료'로 변경
        if status == '읽지않음':
            status = '발송완료'
            
        # D열에 발송시각, E열에 상태 업데이트
        range_name = f'{sheet_name}!D{row_index}:E{row_index}'
        values = [[sent_time, status]]
        body = {'values': values}
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()
    except Exception as e:
        logging.error(f"스프레드시트 업데이트 중 오류 발생: {str(e)}")

def send_email(email_titles, email_contents, user_id, user_pw, recipient_data, selected_db):
    logging.info("\n=== 메일 발송을 시작합니다 ===")
    driver = create_driver()
    keep_browser_open = prevent_browser_close(driver)
    atexit.register(keep_browser_open)
    
    # 스프레드시트 서비스 초기화
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    driver.maximize_window()
    driver.get("https://mail.naver.com/v2/new")

    try:
        # 로그인
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#id")))
        id_element = driver.find_element(By.CSS_SELECTOR, "#id")
        pyperclip.copy(user_id)
        id_element.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        pw_element = driver.find_element(By.CSS_SELECTOR, "#pw")
        pyperclip.copy(user_pw)
        pw_element.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)

        driver.find_element(By.CSS_SELECTOR, ".btn_login").click()
        logging.info("\n=== 로그인이 완료되었습니다 ===")

        # 메일 작성 페이지 대기
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#recipient_input_element")))

        total_recipients = len(recipient_data)
        for idx, recipient in enumerate(recipient_data, 1):
            name, email, row_idx = recipient  # row_idx 추가

            # HTML 모드 선택 (매 메일마다 확실히 선택)
            try:
                html_bt = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.editor_mode_select button[value='HTML']"))
                )
                html_bt.click()
                time.sleep(1)
            except Exception:
                pass

            # 수신자 입력
            address_element = driver.find_element(By.CSS_SELECTOR, "#recipient_input_element")
            address_element.clear()
            pyperclip.copy(email)
            address_element.send_keys(Keys.CONTROL, 'v')
            address_element.send_keys(Keys.ENTER)
            time.sleep(1)

            # 제목 입력 (B1:D1 중 하나만 선택)
            title_element = driver.find_element(By.CSS_SELECTOR, "#subject_title")
            title_element.clear()
            email_title = random.choice(email_titles).replace("{이름}", name)
            pyperclip.copy(email_title)
            title_element.send_keys(Keys.CONTROL, 'v')
            pyautogui.press("tab")
            time.sleep(0.5)

            # 내용 입력 (전체 템플릿)
            email_content = random.choice(email_contents).replace("{이름}", name)
            pyperclip.copy(email_content)
            pyautogui.hotkey("ctrl", "a")  # 기존 내용 전체 선택
            pyautogui.hotkey("ctrl", "v")  # 새 내용으로 덮어쓰기
            time.sleep(2)

            #전송 버튼 클릭
            send_button = driver.find_element(By.CSS_SELECTOR, ".button_write_task")
            send_button.click()
            time.sleep(5)

            # 수신확인함으로 이동
            driver.get("https://mail.naver.com/v2/folders/2")
            time.sleep(5)

            # 메일 항목 정보 가져오기
            try:
                mail_items = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.mail_item.reception"))
                )
                if mail_items:
                    latest_mail = mail_items[0]  # 가장 최근 메일
                    # 이메일 주소만 정확하게 가져오기
                    recipient_element = latest_mail.find_element(By.CSS_SELECTOR, ".recipient_link")
                    recipient = recipient_element.text.split('\n')[-1].strip()  # 마지막 줄의 이메일 주소만 가져오기
                    status = latest_mail.find_element(By.CSS_SELECTOR, ".sent_status").text
                    sent_time = latest_mail.find_element(By.CSS_SELECTOR, ".sent_time").text
                    
                    # 오늘 날짜 추가
                    from datetime import datetime
                    today = datetime.now()
                    sent_time = f"{today.strftime('%y년 %m월 %d일')} {sent_time}"
                    
                    logging.info(f"\n=== 메일 발송 상태 ===")
                    logging.info(f"받는사람: {recipient}")
                    logging.info(f"상태: {status}")
                    logging.info(f"발송시각: {sent_time}")

                    # 이메일 주소 일치 여부 확인
                    if recipient.strip().lower() != email.strip().lower():
                        logging.warning(f"이메일 주소 불일치! 예상: {email}, 실제: {recipient}")
                        status = "미발송 (이메일 불일치)"
                        sent_time = "-"

                    # 스프레드시트에 상태 업데이트 (실제 행 번호 사용)
                    update_sheet_status(service, SPREADSHEET_ID, selected_db, row_idx, sent_time, status)
            except Exception as e:
                # 상태 확인 실패 시 '미발송'으로 기록
                logging.error(f"상태 확인 실패: {str(e)}")
                update_sheet_status(service, SPREADSHEET_ID, selected_db, row_idx, "-", "미발송 (확인실패)")

            logging.info(f"\n=== 진행상황: {idx}/{total_recipients} 완료 ===")
               
            if idx < total_recipients:  # 마지막 메일이 아닌 경우에만 대기
                wait_time = int(random.uniform(3, 70))  # 정수로 변환
                logging.info(f"\n다음 메일 발송까지 대기...")
                
                # 카운트다운 표시
                for remaining in range(wait_time, 0, -1):
                    sys.stdout.write(f"\r{remaining}초 남음...")
                    sys.stdout.flush()
                    time.sleep(1)
                print("\r대기 완료!            ")  # 이전 카운트다운 텍스트 덮어쓰기

                # 새 메일 작성 페이지로 이동
                driver.get("https://mail.naver.com/v2/new")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#recipient_input_element")))
                time.sleep(2)

    except (TimeoutException, NoSuchElementException, WebDriverException) as e:
        logging.error(f"\n=== 오류가 발생했습니다: {str(e)} ===")
    finally:
        logging.info("\n=== 메일 발송이 완료되었습니다 ===")
        logging.info("Enter 키를 눌러 브라우저를 종료하세요...")
        keep_browser_open()
        input()

if __name__ == "__main__":
    try:
        email_titles, email_contents, user_id, user_pw, recipient_data, selected_db = get_data_from_sheets()
        send_email(email_titles, email_contents, user_id, user_pw, recipient_data, selected_db)
    except Exception as e:
        logging.error(f"\n=== 오류가 발생했습니다: {str(e)} ===")