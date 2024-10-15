
#? ----------------------
#? 1012 수정내용
#? 하나의 행에 링크나 댓글 내용이 없으면
#? 검사 대상에서 아에 제외시키도록 로직을 추가함
#? 유효성 스캔 하는 과정에서 구글 api를 지나치게 사용할 수 있기 때문에
#? 0.5초 딜레이를 추가함
#?
#? 구글 폴더 주소를 입력하면 폴더로, 파일을 입력하면 파일 단위로 스캔하도록 수정함
#? ----------------------



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
from googleapiclient.discovery import build
from google.oauth2 import service_account

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


def clean_text_with_br_preservation(element):
    # Element를 문자열로 변환한 후에 <br>을 \n으로 변환
    html_content = str(element)
    # <br> 태그를 줄바꿈으로 변환
    html_content = html_content.replace('<br>', ' ').replace('<br/>', ' ').replace('</br>', ' ')
    # BeautifulSoup을 사용하여 HTML 파싱 및 텍스트 추출
    text = BeautifulSoup(html_content, "html.parser").get_text()
    text2 = clean_text(text.strip())
    return text2





def get_gdrive_service():
    creds = service_account.Credentials.from_service_account_file('foroncomm-57ce18a35975.json')
    scoped_creds = creds.with_scopes(['https://www.googleapis.com/auth/drive.readonly'])
    return build('drive', 'v3', credentials=scoped_creds)

def list_sheets_in_folder(service, folder_id):
    results = service.files().list(q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet'",
                                   fields="files(id, name)").execute()
    return results.get('files', [])

def process_sheet(sheet, driver, sheet_name):
    print(f"파일 로드 중: '{sheet_name}'")
    print("--------------------------------")
    
    rows = sheet.get_all_values()[3:]  # 4번째 행부터 데이터 시작
    total_rows = len(rows)
    batch_size = 20  # 한 번에 처리할 타겟 수
    total_batches = (total_rows // batch_size) + (1 if total_rows % batch_size != 0 else 0)

    for batch_num in range(total_batches):
        print(f"\n'{sheet_name}' 시트의 {batch_num + 1}/{total_batches}번째 실행을 시작합니다.")
        print("--------------------------------")

        targets = []
        start_idx = batch_num * batch_size
        end_idx = min(start_idx + batch_size, total_rows)

        print("----배치 전체 유효성 스캔을 시작합니다----")
        for idx, row in enumerate(rows[start_idx:end_idx], start=start_idx + 4):

            
            if len(row) >= 12 and not row[11].strip():  # L열이 비어 있는지 확인
                target_link = row[4] if len(row) > 4 else ""
                target_text = clean_text(row[5]) if len(row) > 5 else ""

                if not target_link or not target_text:
                    print(f"Row {idx}: 링크나 타겟 텍스트가 없습니다.")
                    sheet.update_cell(idx, 12, "데이터 부족")
                    time.sleep(0.5)
                    continue


                # 모바일 링크를 www로 변경
                if target_link.startswith('http://m.'):
                    target_link = target_link.replace('http://m.', 'http://')
                elif target_link.startswith('https://m.'):
                    target_link = target_link.replace('https://m.', 'https://')



                if not target_link.startswith('http'):
                    print(f"Row {idx}: 유효하지 않은 링크 형식입니다. 링크: '{target_link}'")
                    sheet.update_cell(idx, 12, "확인불가")
                    time.sleep(0.5)
                    continue

                if not target_text:
                    print(f"Row {idx}: 타겟 텍스트가 비어 있습니다.")
                    sheet.update_cell(idx, 12, "확인불가")
                    time.sleep(0.5)
                    continue
                

                targets.append((idx, target_link, target_text))

        print("----배치 전체 유효성 스캔 완료----")
        
        if not targets:
            print(f"{sheet_name} 시트의 {batch_num + 1}번째 실행에서 비어있는 행을 찾을 수 없습니다.")
            continue
        

        


        # 타겟 링크와 타겟 텍스트 출력 및 처리
        for i, (row_idx, link, text) in enumerate(targets, start=1):
            print(f"Target {i}: Row {row_idx}, {text}")

            driver.get(link)
            time.sleep(1)  # 페이지 로드 대기

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
                target_comment_time = None
                target_comment_index = None

                for idx, comment in enumerate(comments):
                    comment_text_view = comment.find('span', class_='text_comment')
                    comment_time_view = comment.find('span', class_='comment_info_date')

                    if comment_text_view and comment_time_view:

                        comment_text = clean_text_with_br_preservation(comment_text_view)
                        comment_time = comment_time_view.text.strip()  # 댓글 작성 시간 가져오기
                        print(f"댓글 {idx + 1}: {comment_text}, 작성 시간: {comment_time}")

                        if text in comment_text:
                            target_comment_time = comment_time
                            target_comment_index = idx + 1
                            comment_found = True

                            # 다음 댓글 시간 가져오기
                            if idx + 1 < len(comments):
                                next_comment_time_view = comments[idx + 1].find('span', class_='comment_info_date')
                                if next_comment_time_view:
                                    next_comment_time = next_comment_time_view.text.strip()
                                    if target_comment_time > next_comment_time:
                                        sheet.update_cell(row_idx, 12, f"{target_comment_index} (시간 밀림)")
                                    else:
                                        sheet.update_cell(row_idx, 12, target_comment_index)
                            else:
                                sheet.update_cell(row_idx, 12, target_comment_index)
                            break

            if not comment_found:
                print(f'Row {row_idx}: 댓글을 찾을 수 없습니다.')
                sheet.update_cell(row_idx, 12, "없음")

            print('-------------------------------')
            time.sleep(2)  # 다음 타겟 링크로 이동 전 딜레이 추가

        print(f"'\n{sheet_name}' 시트의 {batch_num + 1}번째 실행의 처리가 완료되었습니다. \n30초간 대기 후 다음 실행이 이어집니다.")
        print("--------------------------------------------------------")
        print("--------------------------------------------------------")
        time.sleep(30)

    print(f"{sheet_name} 시트의 모든 데이터 처리가 완료되었습니다.\n")

def main():
    try:
        # 사용자로부터 Google Drive 파일 또는 폴더 URL 입력받기
        url = input("구글 드라이브 파일 또는 폴더 URL을 입력하세요: ")

        # URL에서 ID 추출
        if "folders" in url:
            # 폴더 URL에서 폴더 ID 추출
            folder_id = url.split('/')[-1]
            is_folder = True
        elif "spreadsheets" in url:
            # 파일 URL에서 파일 ID 추출
            file_id = url.split('/d/')[1].split('/')[0]
            is_folder = False
        else:
            print("올바른 구글 드라이브 파일 또는 폴더 URL을 입력하세요.")
            return

        # Google Drive API 서비스 생성
        service = get_gdrive_service()

        if is_folder:
            # 폴더 내 모든 구글 시트 파일 목록 가져오기
            sheets = list_sheets_in_folder(service, folder_id)

            if not sheets:
                print("폴더에 구글 시트 파일이 없습니다.")
                return

        else:
            # 파일을 바로 처리
            sheets = [{'id': file_id, 'name': '입력된 파일'}]

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
            
            # 각 구글 시트에 대해 작업 수행
            for sheet_info in sheets:
                sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_info['id']}/edit"
                print("--------------------------------------------------------")
                print(f"'{sheet_info['name']}' 시트에 대해 작업을 시작합니다.")
                
                # Google Sheets API 설정
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                creds = ServiceAccountCredentials.from_json_keyfile_name('foroncomm-57ce18a35975.json', scope)
                client = gspread.authorize(creds)

                # Google Sheets 열기
                sheet = client.open_by_url(sheet_url).sheet1  # 첫 번째 시트를 열기

                # 파일 로드 및 작업 수행
                process_sheet(sheet, driver, sheet_info['name'])
                print(f"{sheet_info['name']} 시트에 대한 작업이 완료되었습니다.")
                print('-------------------------------')

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
