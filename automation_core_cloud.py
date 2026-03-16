import logging
import time
import re
import os
import json
import math
import ctypes
from datetime import datetime
from PyPDF2 import PdfReader
from playwright.sync_api import sync_playwright

logging.basicConfig(level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename='process_error.log',
                    encoding='utf-8')

IGNORED_FILES = set()
CONFIG_PATH = "config.json"
CRO_DB_PATH = "cro_contacts.json"

selector_map = {
    "item_22323_52866_0_0": "#item_22360_2028638_0_0", "item_22323_52865_0_0": "#item_22360_2028639_0_0",
    "item_22323_52864_0_0": "#item_22360_2028639_0_0", "item_22323_291222_0_0": "#item_22360_2028643_0_0",
    "item_22323_4383463_0_0": "#item_22360_4383461_0_0", "item_22323_4383464_0_0": "#item_22360_4383462_0_0",
    "item_22323_2028988_0_0": "#item_22360_2028989_0_0", "item_22323_2028990_0_0": "#item_22360_2028989_0_0",
    "item_22323_63689_0_0": "#item_22360_2028644_0_0", "item_22323_689306_0_0": "#item_22360_2028645_0_0",
    "item_22323_63690_0_0": "#item_22360_2028646_0_0", "item_22323_530313_0_0": "#item_22360_2028651_0_0",
    "item_22323_2881324_0_0": "#item_22360_2891005_0_0", "item_22323_3396617_0_0": "#item_22360_3396624_0_0",
    "item_22323_2028991_0_0": "#item_22360_2028652_0_0",
}

def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"⚠️ JSON 읽기 실패 ({path}): {e}")
    return default

def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        logging.error(f"⚠️ JSON 저장 실패 ({path}): {e}")

def load_config():
    return load_json(CONFIG_PATH, {
        "site_url": "https://drn.cubecdms.com/login",
        "id": "jslee@crscube.io",
        "pw": "cube0216!@"
    })

def save_config(data):
    save_json(CONFIG_PATH, data)

def extract_pdf_text(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        return "".join([page.extract_text() for page in reader.pages])
    except Exception as e:
        logging.error(f"⚠️ PDF 텍스트 추출 실패: {e}")
        return None

def parse_pdf_data(text):
    data = {}
    patterns = {
        'protocol_title': r"임상연구제목\s*[:：\s]*((?:.|\n)+?)\s*·?\s*Protocol No\.?",
        'protocol_no': r"Protocol No\.?\s*[:：\s]*([^\n\r]+)",
        'sponsor': r"Sponsor\s*[:：,\s]*([^\n\r]+)",
        'phase': r"Phase\s*[:：\s]*([^\n\r]+)",
        'duration_month': r"예상 운용 기간\s*[:：\s]*(\d+)\s*개월",
        'subject_num': r"시험대상자수\s*[:：\s]*(\d+)",
        'monthly_fee': r"매월\s*([\d,]+)\s*원이다",
        'billing_cycle': r"매\s*(\d+)\s*개월마다\s*전자세금계산서"
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL | re.MULTILINE)
        if match:
            val = match.group(1).strip()
            if key == 'protocol_no': val = re.sub(r'\s*-\s*', '-', val)
            elif key in ['monthly_fee', 'billing_cycle']: val = val.replace(",", "")
            data[key] = val
        else: data[key] = ""
    
    cro_match = re.search(r'([^\s]+)\s*\(“고객”\)', text)
    if cro_match:
        cro = re.sub(r'\(주\)|㈜|주식회사', '', cro_match.group(1)).strip()
        if "씨엔알리서치" in cro: cro = "CNR"
        data['cro_name'] = cro

    checked_symbols, unchecked_symbol = r'[☑☒]', r'☐'
    pdf_checkbox_mapping = {
        'item_22323_52866_0_0': r'cubeCDMS',
        'item_22323_52865_0_0': r'cube[Il]WRS®(?!\s*\(Random)', 
        'item_22323_52864_0_0': r'cube[Il]WRS®\s*\(Random',
        'item_22323_291222_0_0': r'cubePRO',
        'item_22323_63689_0_0': r'cubeSAFETY',
        'item_22323_689306_0_0': r'cubeTMF',
        'item_22323_63690_0_0': r'cubeCTMS',
        'item_22323_530313_0_0': r'cubeConsent',
        'item_22323_2881324_0_0': r'cubeDDC',
        'item_22323_3396617_0_0': r'cubeRBQM',
        'item_22323_2028988_0_0': r'File\s*server\s*computer\s*\(독립적\s*평가자',
        'item_22323_2028990_0_0': r'File\s*server\s*computer\s*\(file\s*upload\s*only'
    }
    selected_ids = []
    for cb_id, pattern in pdf_checkbox_mapping.items():
        search_regex = f'{checked_symbols}[^{unchecked_symbol}☑☒]*{pattern}'
        if re.search(search_regex, text, re.IGNORECASE): selected_ids.append(cb_id)
    data['selected_checkbox_ids'] = selected_ids
    data['contract_year'] = str(datetime.now().year)
    return data

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", name.replace("(주)", "").replace("㈜", "")).strip().replace(" ", "_")

def rename_files_based_on_data(current_pdf_path, data):
    try:
        folder = os.path.dirname(current_pdf_path)
        current_docx = current_pdf_path.replace(".pdf", ".docx")
        sponsor, protocol = data.get('sponsor', ''), data.get('protocol_no', '')
        if not sponsor or not protocol: return current_pdf_path
        new_base = f"WO{data['contract_year']}_{sanitize_filename(sponsor)}_{sanitize_filename(protocol)}"
        new_pdf = os.path.join(folder, f"{new_base}.pdf")
        new_docx = os.path.join(folder, f"{new_base}.docx")
        if os.path.exists(current_pdf_path) and current_pdf_path != new_pdf:
            if os.path.exists(new_pdf): os.remove(new_pdf)
            os.rename(current_pdf_path, new_pdf); IGNORED_FILES.add(os.path.abspath(new_pdf)) 
            current_pdf_path = new_pdf
        if os.path.exists(current_docx) and current_docx != new_docx:
            if os.path.exists(new_docx): os.remove(new_docx)
            os.rename(current_docx, new_docx); IGNORED_FILES.add(os.path.abspath(new_docx)) 
        return current_pdf_path
    except Exception as e:
        logging.error(f"파일 이름 변경 실패: {e}")
        return current_pdf_path

def input_calculated_amounts(page, total_billing_count, form_data):
    checkbox_amounts = form_data["checkbox_amounts"]
    total_amount = form_data["total_amount"]
    selector_base_map_js = "{\n" + ",\n".join([f'"{k}": "{v}"' for k, v in selector_map.items()]) + "\n}"
    js_script = f"""
    ({{ count, checkboxAmounts, totalAmount }}) => {{
        const selectorBaseMap = {selector_base_map_js};
        const setValue = (selector, value) => {{
            const el = document.querySelector(selector);
            if(el){{ el.value = value; el.dispatchEvent(new Event('input', {{bubbles:true}})); el.dispatchEvent(new Event('change', {{bubbles:true}})); }}
        }};
        setTimeout(() => {{
            for(let i=1; i<=count; i++){{
                for(const [cbId, amount] of Object.entries(checkboxAmounts)){{
                    if(amount <= 0) continue;
                    const base = selectorBaseMap[cbId];
                    if(base) setValue(base.replace(/_0_\\d+$/, `_0_${{i}}`), amount);
                }}
                setValue(`#item_22360_52993_0_${{i}}`, totalAmount);
            }}
        }}, 500);
    }}
    """
    page.evaluate(js_script, {"count": total_billing_count, "checkboxAmounts": checkbox_amounts, "totalAmount": total_amount})

def safe_click(page, sel, desc="", timeout=10000, ask_user_on_fail=True):
    try:
        page.wait_for_selector(sel, state="visible", timeout=timeout)
        time.sleep(0.5)
        page.click(sel)
        print(f"✅ {desc} 클릭")
        return True
    except Exception as e:
        msg = f"⛔ {desc} 클릭 실패: {e}"
        print(msg)
        logging.error(msg)
        if ask_user_on_fail:
            # 안전한 Windows 네이티브 MessageBox 호출 (스레드 충돌 방지)
            MB_YESNO = 4
            MB_ICONWARNING = 0x30
            MB_TOPMOST = 0x40000
            IDYES = 6
            res = ctypes.windll.user32.MessageBoxW(
                0, 
                f"'{desc}' 작업을 자동 수행하는데 실패했습니다.\n\n수동으로 브라우저에서 조작을 완료하셨다면 [예(Y)]를 눌러 계속 진행하고,\n이 단계를 건너뛰려면 [아니요(N)]를 누르세요.", 
                "수동 조작 대기", 
                MB_YESNO | MB_ICONWARNING | MB_TOPMOST
            )
            if res == IDYES:
                return True
            return False
        return False

def fill_cro_contacts(page, cro_info):
    person = (cro_info.get("person") or "").strip()
    phone = (cro_info.get("phone") or "").strip()
    email = (cro_info.get("email") or "").strip()
    if person:
        try: page.fill("#item_58699_133829_0_0", person)
        except Exception as e: logging.error(f"담당자 이름 입력 실패: {e}")
    if phone:
        try: page.locator('input[name="item_58700_133830_0_0"]').fill(phone)
        except Exception as e: logging.error(f"전화번호 입력 실패: {e}")
    if email:
        try: page.fill("#item_58701_133831_0_0", email)
        except Exception as e: logging.error(f"이메일 입력 실패: {e}")

def run_automation(res, status_callback=None):
    login, form = res["login_data"], res["form_data"]

    def log_status(msg):
        print(msg)
        if status_callback:
            status_callback(msg)

    with sync_playwright() as p:
        try:
            log_status("🌐 Playwright 브라우저 연결 시도 중...")
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = None
            for pg in context.pages:
                if "cubeManager" in pg.title(): page = pg; break
            if not page: page = context.new_page()

            max_retries = 3
            retry_count = 0
            subject_btn_found = False

            while retry_count < max_retries and not subject_btn_found:
                try:
                    log_status(f"🔄 사이트 접속 시도 ({retry_count + 1}/{max_retries})...")
                    page.goto(login['login_url'])
                    page.wait_for_selector("#txt_user_id", state="visible", timeout=15000)
                    page.fill("#txt_user_id", login['user_id'])
                    page.fill("#txt_password", login['user_pw'])
                    page.click("#btn_login")
                    try: page.click("#btn_change_login", timeout=1000)
                    except: pass
                    
                    page.wait_for_selector("body", timeout=30000)
                    page.wait_for_selector('a[title="subject"] > div.l-menu-subject-off', state="visible", timeout=15000)
                    subject_btn_found = True
                except Exception as loop_e:
                    retry_count += 1
                    logging.error(f"초기 진입 지연: {loop_e}")
                    if retry_count >= max_retries:
                        raise Exception("최대 재시도 횟수를 초과하여 로그인 및 초기 화면 진입에 실패했습니다.")

            log_status("✅ 기본 화면 진입 및 과제 추가 진행 중...")
            safe_click(page, 'a[title="subject"] > div.l-menu-subject-off', "메뉴: subject", ask_user_on_fail=False) 
            page.wait_for_selector("body")
            safe_click(page, "#btn_add_subject", "과제 추가 버튼", ask_user_on_fail=False)
            safe_click(page, "#sbj_team-25496", "팀 선택")
            safe_click(page, 'input[type="button"].green.ui-button', "팀 확정")
            
            safe_click(page, "#item_22308_52843_0_0", "날짜 팝업 오픈")
            safe_click(page, f'//table[contains(@class,"ui-datepicker-calendar")]//a[text()="{datetime.now().day}"]', "오늘 날짜 클릭")
            try: page.click('//button[contains(text(),"닫기")]', timeout=1000)
            except: pass

            safe_click(page, "#item_1637533_4371919_0_0-1", "기본체크 1") 
            safe_click(page, "#item_1637533_4568964_0_0-1", "기본체크 2") 
            safe_click(page, "#item_58697_133827_0_0-2", "기본체크 3")
            fill_cro_contacts(page, res["cro_info"])

            safe_click(page, 'div.placeholder:text("[NotSelected]")', "국가 선택") 
            safe_click(page, '//div[@class="pt5"]/span[contains(text(),"KR:한국")]', "한국(KR) 선택")
            safe_click(page, "#btn_add_next", "다음 버튼")
            page.get_by_role("button", name="확인").click()

            page.wait_for_selector("#item_22311_52851_0_0", state="visible", timeout=10000)
            page.fill("#item_22311_52851_0_0", datetime.now().strftime("%Y-%m-%d"))
            try: page.click('//button[contains(text(),"닫기")]', timeout=1000)
            except: pass

            log_status("📝 과제 정보 입력 중...")
            page.fill("#item_22312_52852_0_0", form['protocol_no']); page.fill("#item_22313_52853_0_0", form['protocol_title'])
            safe_click(page, 'div.placeholder:text("[NotSelected]")', "Phase 선택 오픈")
            phase = form['phase']
            if phase == "BE/BA": safe_click(page, '//div[contains(@class,"pt5")]/span[normalize-space(text())="BE"]', f"Phase: BE")
            elif "&" in phase:
                for part in [p.strip() for p in phase.split("&")]:
                    if safe_click(page, f'//div[contains(@class,"pt5")]/span[contains(text(),"{part}")]', rf"Phase: {part}"):
                        break
            else: safe_click(page, f'//div[contains(@class,"pt5")]/span[contains(text(),"{phase}")]', f"Phase: {phase}")

            page.fill("#item_35653_52857_0_0", str(form['duration_month'])); page.fill("#item_22318_52860_0_0", str(form['subject_num']))
            if res["selected_checkbox_ids"]:
                for cb in res["selected_checkbox_ids"]:
                    try: page.click(f"#{cb}")
                    except Exception: pass
            
            safe_click(page, '#item_22322_52863_0_0-2', "추가사항 1") 
            safe_click(page, '#item_223157_531110_0_0-2', "추가사항 2") 
            safe_click(page, '#btn_save_b', "1차 저장 버튼")
            
            log_status("💰 비용청구 탭 설정 중...")
            safe_click(page, 'div.crf-schedule-visit:text("All")', "All 탭") 
            safe_click(page, 'div.left[title="비용청구"]', "비용청구 탭")
            page.wait_for_load_state("networkidle"); time.sleep(2)

            count = math.ceil(form['duration_month'] / form['billing_cycle'])
            add_btn = 'div.icon[style*="padding-top: 40px"][style*="padding-bottom: 41px"] > a.mr5[data-role="add"][data-mode="edit"][title="추가"]'
            for _ in range(count):
                try: 
                    page.wait_for_selector(add_btn, state="visible", timeout=10000)
                    page.click(add_btn)
                    time.sleep(0.3)
                except Exception: pass
            
            time.sleep(3)
            try:
                last_sel = f"select#item_22360_52996_0_{count + 1}"
                page.wait_for_selector(last_sel, state="attached", timeout=10000)
                js_loop_script = "count => { for(let i=1; i<=count; i++){ const sel = `select#item_22360_52996_0_${i}`; const el = document.querySelector(sel); if(el){ el.value = '7'; el.dispatchEvent(new Event('change', {bubbles:true})); } } }"
                page.evaluate(js_loop_script, count + 1)
            except Exception as e: logging.warning(f"비용청구 select 값 설정 오류: {e}")

            page.wait_for_load_state("networkidle"); time.sleep(3)
            input_calculated_amounts(page, count + 1, form); time.sleep(3)

            try: page.click("#item_460278_1156461_0_0-1")
            except: pass
            try: page.click("#item_460278_1927884_0_0-1")
            except: pass
            try: page.fill("#item_460278_1156460_0_0", str(form['monthly_fee']))
            except: pass

            cycle_map = {1: "-1", 3: "-2", 6: "-3"}
            if form['billing_cycle'] in cycle_map:
                tgt = f"#item_460278_1155912_0_0{cycle_map[form['billing_cycle']]}"
                safe_click(page, tgt, "청구 주기 선택")
            
            log_status("💾 폼 저장 (완료 전 단계체크 방식을 통해 확인 요망)")
            safe_click(page, 'a#btn_save_b.button.green[name="btn_save"]', "최종 저장 버튼")
            
            try:
                page.wait_for_selector('div#popover', state="visible", timeout=5000)
                page.click('input[type="radio"][id="modify_reason-1"]'); page.click('input[type="button"][value="저장"]')
                page.wait_for_selector('input[type="button"][value="확인"]', state="visible", timeout=5000); page.click('input[type="button"][value="확인"]')
            except: pass
            
            log_status("✅ 모든 작업 완료!")
            return {"success": True, "message": "성공적으로 입력 작업을 완료했습니다."}

        except Exception as e:
            logging.error(f"프로세스 진행 중 오류: {e}", exc_info=True)
            return {"success": False, "message": str(e)}
