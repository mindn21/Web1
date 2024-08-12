import os
import time
import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import sys
import traceback
import unicodedata
import re

def clean_text(text):
    # 유니코드 정규화
    text = unicodedata.normalize('NFC', text)
    # HTML 엔티티 제거
    text = BeautifulSoup(text, "html.parser").text
    # 공백과 줄바꿈 처리
    text = ' '.join(text.split())
    # 특수 문자 제거 및 소문자 변환
    text = re.sub(r'[^\w\s]', '', text).lower()
    return text

def main():
    try:
        # 사용자로부터 Google Sheets URL 입력받기
        sheet_url = input("구글 시트 URL을 입력하세요: ")

        # sheet_url = "https://docs.google.com/spreadsheets/d/1YdXhz-gwEShBeluBTU63Gjer7XfblsAMV4TVELC4D9o/edit?gid=1975501167#gid=1975501167"

        # Google Sheets API 설정
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('foroncomm-57ce18a35975.json', scope)  # 인증 JSON 파일의 경로
        client = gspread.authorize(creds)

        # Google Sheets 열기
        sheet = client.open_by_url(sheet_url).sheet1  # 첫 번째 시트를 열기

        targets = []
        count = 20 #몇개의 타겟 링크와 타겟 텍스트 수집할지

        for idx, row in enumerate(sheet.get_all_values()[3:], start=4):  # 4번째 행부터 데이터 시작
            if len(row) >= 12:
                if not row[11].strip():  # L열이 비어 있는지 확인
                    target_link = row[4]  # E열의 값 (인덱스 4)
                    target_text = clean_text(row[5])  # F열의 값 (인덱스 5)
                    targets.append((idx, target_link, target_text))
                    if len(targets) >= count:  
                        break
            else:
                # L열이 존재하지 않는 경우에도 처리
                target_link = row[4] if len(row) > 4 else ""
                target_text = clean_text(row[5]) if len(row) > 5 else ""
                targets.append((idx, target_link, target_text))
                if len(targets) >= count:
                    break

        if not targets:
            print("L열이 비어있는 행을 찾을 수 없습니다.")
            sys.exit()

        # 타겟 링크와 타겟 텍스트 출력
        for i, (row_idx, link, text) in enumerate(targets, start=1):
            print(f"Target {i}: Row {row_idx}, Link: {link}, Text: {text}")

        # 로그인 수행
        shortcut_path = os.path.abspath('./chrome.lnk')
        os.startfile(shortcut_path)
        time.sleep(2)

        options = webdriver.ChromeOptions()
        options.debugger_address = "localhost:9222"

        driver_path = './chromedriver.exe' 
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)

        print('-------------------------------')
        print('-------------------------------')
        print("로그인 페이지로 이동합니다.")
        driver.get('https://nid.naver.com/nidlogin.login')

        print("사용자가 직접 로그인하도록 기다립니다.")

        def show_message_box():
            root = tk.Tk()
            root.withdraw()
            return messagebox.askyesno("로그인 확인", "수동으로 로그인을 완료한 후 '예'를 클릭하세요.\n종료하려면 '아니요'를 클릭하세요.")

        user_response = show_message_box()

        if user_response and "https://www.naver.com/" in driver.current_url:
            print("로그인 성공, 네이버 메인 페이지로 이동했습니다.")
            print("타겟 링크로 이동합니다")
            print('-------------------------------')
            
            for row_idx, target_link, target_text in targets:
                driver.get(target_link)
                time.sleep(1)  # 페이지 로드 대기

                # 페이지가 iframe 내부에 있는지 확인하고, 있다면 해당 iframe으로 전환
                try:
                    driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, 'iframe#cafe_main'))
                    print("iframe으로 전환 성공")
                    time.sleep(1)
                except Exception as e:
                    print(f"iframe으로 전환하는 중 오류 발생: {e}")
                    print("게시글이 삭제되었습니다.")
                    time.sleep(1)
                    sheet.update_cell(row_idx, 12, "게시글 삭제")
                    print('-------------------------------')
                    continue  # 다음 타겟 링크로 이동

                # 댓글 찾기
                page_html = driver.page_source
                soup = BeautifulSoup(page_html, 'html.parser')

                comments_list = soup.find('ul', class_='comment_list')
                comment_found = False  # comment_found 변수 초기화

                if comments_list:
                    comments = comments_list.find_all('li', class_='CommentItem')
                    for idx, comment in enumerate(comments):
                        comment_text_view = comment.find('span', class_='text_comment')
                        if comment_text_view:
                            comment_text = clean_text(comment_text_view.text.strip())
                            print(f"댓글 {idx + 1}: {comment_text}")  # 각 댓글의 내용을 출력
                            if target_text in comment_text:
                                print(f'Row {row_idx}은(는) {idx + 1}번째 댓글입니다.')
                                sheet.update_cell(row_idx, 12, idx + 1)  # L열 업데이트
                                comment_found = True
                                break

                if not comment_found:
                    print(f'Row {row_idx}: 댓글을 찾을 수 없습니다.')
                    sheet.update_cell(row_idx, 12, "없음")

                print('-------------------------------')
                time.sleep(2)  # 새로운 순위를 확인하기 전 딜레이 추가

        else:
            print("로그인에 실패했거나 사용자가 '아니요'를 선택했습니다. 프로그램을 종료합니다.")
            driver.quit()
            sys.exit()

        # 브라우저 종료
        print('-------------------------------')
        driver.quit()

    except Exception as e:
        print(f"오류 발생: {e}")
        traceback.print_exc()  # 에러 발생 시 자세한 정보와 줄 번호 출력

    finally:
        # 프로그램 종료 시, 엔터 키 입력 대기
        input("프로그램이 종료되었습니다. 엔터 키를 누르면 닫힙니다.")

if __name__ == "__main__":
    main()
