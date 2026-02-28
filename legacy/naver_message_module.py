import random
import logging
import sys
from pathlib import Path
from googleapiclient.discovery import build
from auth import get_credentials
from datetime import datetime
import calendar

# module_notion 경로를 파이썬 경로에 추가
module_notion_path = Path("~/Documents/github_cloud/module_notion").expanduser()
if str(module_notion_path) not in sys.path:
    sys.path.insert(0, str(module_notion_path))

from notion_reader import get_database_items, extract_page_id_from_url, print_database_items

def get_week_number(date):
    """
    주어진 날짜의 주차를 계산하는 함수 (2025년 기준)
    
    Args:
        date (datetime): 날짜
        
    Returns:
        int: 주차 (1-5)
    """
    # 해당 월의 첫 날
    first_day = date.replace(day=1)
    # 첫 날의 요일 (0: 일요일, 6: 토요일)
    first_day_weekday = first_day.weekday() + 1  # 월요일(0)을 일요일(0)로 변경
    if first_day_weekday == 7:  # 토요일인 경우
        first_day_weekday = 0
        
    # 해당 날짜가 몇 번째 주인지 계산
    week_number = ((date.day + first_day_weekday - 1) // 7) + 1
    return week_number

class InstagramMessageTemplate:
    def __init__(self, template_sheet_id, template_sheet_name):
        self.template_sheet_id = template_sheet_id
        self.template_sheet_name = template_sheet_name

    def get_message_templates(self):
        """
        구글 스프레드시트에서 메시지 템플릿을 가져와 조합하는 함수
        
        Returns:
            list: 조합된 메시지 템플릿 리스트
        """
        logging.info("메시지 템플릿 가져오기 시작")
        try:
            creds = get_credentials()
            service = build('sheets', 'v4', credentials=creds)

            sheet = service.spreadsheets()
            # 본문 템플릿만 가져오기 (B2:D4 범위)
            result = sheet.values().get(spreadsheetId=self.template_sheet_id,
                                        range=f'{self.template_sheet_name}!B2:D4').execute()
            values = result.get('values', [])

            if not values or len(values) < 3:  # 3개의 행(인사말, 제안, 맺음말)이 필요
                logging.warning('메시지 템플릿을 찾을 수 없습니다.')
                return ["안녕하세요"]

            # 각 파트별로 무작위 선택 (제목 제외)
            greeting = random.choice(values[0]) if values[0] else ""  # B2:D2 (인사말)
            proposal = random.choice(values[1]) if values[1] else ""  # B3:D3 (제안)
            closing = random.choice(values[2]) if values[2] else ""   # B4:D4 (맺음말)

            # 선택된 파트들을 조합하여 하나의 메시지 생성
            message = f"{greeting}\n\n{proposal}\n\n{closing}"
            
            return [message]  # 리스트 형태로 반환하여 기존 코드와의 호환성 유지

        except Exception as e:
            logging.error(f"메시지 템플릿을 가져오는 중 오류 발생: {e}")
            return ["안녕하세요"]

    def format_message(self, template, name="", notion_list="", total_list=""):
        """
        템플릿의 변수를 실제 값으로 대체하는 함수
        
        Args:
            template (str): 메시지 템플릿
            name (str): 인플루언서 이름
            notion_list (str): 노션 리스트
            total_list (str): 전체 리스트
            
        Returns:
            str: 변수가 대체된 메시지
        """
        # 현재 날짜 기준으로 주차 계산
        now = datetime.now()
        week_number = get_week_number(now)
        week_date = f"{now.month}월 {week_number}주차"
        
        return template.replace("{이름}", name)\
                      .replace("{노션리스트}", notion_list)\
                      .replace("{전체리스트}", total_list)\
                      .replace("{상품리스트}", notion_list)\
                      .replace("{주날짜}", week_date)  # 주날짜 변수 추가 