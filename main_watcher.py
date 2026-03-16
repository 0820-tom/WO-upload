import time
import os
import json
import psutil
import pythoncom
import win32com.client
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import automation_core
import sys

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)

# app.py와 공유할 파싱 결과 파일 경로
PARSED_DATA_PATH = "parsed_data.json"

class DebugHandler(FileSystemEventHandler):
    def on_any_event(self, event):
        if event.event_type == 'created':
            self.check_and_process(event.src_path, "생성됨")
        elif event.event_type == 'moved':
            self.check_and_process(event.dest_path, "이름변경됨")

    def check_and_process(self, filepath, action_type):
        filename = os.path.basename(filepath)
        
        # 임시 파일 무시
        if filename.startswith('~$') or filename.endswith('.tmp') or filename.endswith('.crdownload'):
            return

        # 방금 automation_core가 이름을 바꾼 파일인가? (무시 리스트 체크)
        abs_path = os.path.abspath(filepath)
        if abs_path in automation_core.IGNORED_FILES:
            automation_core.IGNORED_FILES.remove(abs_path)
            return

        if filepath.endswith('.docx'):
            print(f"\n⚡ [감지 성공] 파일이 {action_type}: {filename}")
            self.process_workflow(filepath)

    def process_workflow(self, docx_path):
        print(f"   --> ⏳ 안정화 대기 중...")
        if not self.wait_for_file_ready(docx_path):
            print("   ❌ [실패] 파일 접근 불가")
            return

        pdf_path = docx_path.replace(".docx", ".pdf")
        print(f"   --> 🔄 PDF 변환 시도: {os.path.basename(docx_path)}")
        
        automation_core.IGNORED_FILES.add(os.path.abspath(pdf_path))

        if self.convert_to_pdf_safe(docx_path, pdf_path):
            print("   --> 📄 데이터 파싱 중...")
            try:
                pdf_text = automation_core.extract_pdf_text(pdf_path)
                print(f"   --> 📝 텍스트 추출 완료 (길이: {len(pdf_text) if pdf_text else 0})")

                data = automation_core.parse_pdf_data(pdf_text)
                print(f"   --> 🔎 파싱 결과:")
                print(f"       sponsor      : {data.get('sponsor', '(없음)')}")
                print(f"       protocol_no  : {data.get('protocol_no', '(없음)')}")
                print(f"       protocol_title: {data.get('protocol_title', '(없음)')[:30] if data.get('protocol_title') else '(없음)'}")
                print(f"       cro_name     : {data.get('cro_name', '(없음)')}")

                pdf_path = automation_core.rename_files_based_on_data(pdf_path, data)
                print(f"   --> 📁 파일명 변경 후 경로: {pdf_path}")

                # 파싱 결과 + pdf 경로를 json으로 저장 → app.py가 감지
                data["_auto_pdf_path"] = pdf_path
                data["_timestamp"] = time.time()
                with open(PARSED_DATA_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)

                print(f"   ✅ [완료] 파싱 결과 저장 → app.py로 전달\n")
            except Exception as e:
                import traceback
                print(f"   🔥 [에러] {e}")
                traceback.print_exc()
        else:
            print("   🔥 [에러] PDF 변환 실패")

    def wait_for_file_ready(self, filepath, timeout=30):
        start_time = time.time()
        last_size = -1
        while time.time() - start_time < timeout:
            try:
                if not os.path.exists(filepath): time.sleep(1); continue
                current_size = os.path.getsize(filepath)
                if current_size > 0 and current_size == last_size: return True
                last_size = current_size
                time.sleep(1)
            except OSError: time.sleep(1)
        return False

    def convert_to_pdf_safe(self, docx_path, pdf_path):
        word = None
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(docx_path, ReadOnly=True, Visible=False)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close(SaveChanges=False)
            return True
        except Exception as e:
            print(f"   --> ❌ Word 에러: {e}")
            return False
        finally:
            if word:
                try: word.Quit()
                except: pass
            del word
            pythoncom.CoUninitialize()

def kill_zombie_word():
    print("🧹 Word 프로세스 정리 중...")
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] == "WINWORD.EXE": proc.kill()
        except: pass

if __name__ == "__main__":
    os.system('chcp 65001 > nul')
    kill_zombie_word()
    
    path_to_watch = os.path.join(os.path.expanduser("~"), "Downloads")

    if not os.path.exists(path_to_watch):
        print(f"❌ [경로 오류] 폴더 없음: {path_to_watch}")
    else:
        event_handler = DebugHandler()
        observer = Observer()
        observer.schedule(event_handler, path_to_watch, recursive=False)
        print(f"\n👀 [감시 재시작] 대상: {path_to_watch}\n")
        observer.start()
        try:
            while True: time.sleep(1)
        except KeyboardInterrupt:
            observer.stop()
        observer.join()
