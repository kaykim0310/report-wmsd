import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import json
import sqlite3
import os
import time

# PDF 관련 imports (선택사항)
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

st.set_page_config(layout="wide", page_title="근골격계 유해요인조사")

# 데이터베이스 초기화
def init_db():
    conn = sqlite3.connect('musculoskeletal_survey.db')
    conn.execute("PRAGMA journal_mode=WAL")  # 동시성 개선
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS survey_data
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  session_id TEXT UNIQUE,
                  workplace TEXT,
                  data TEXT,
                  created_at TIMESTAMP,
                  updated_at TIMESTAMP)''')
    conn.commit()
    conn.close()

# 최적화된 데이터 저장 함수
def save_to_db(session_id, data, workplace=None):
    conn = sqlite3.connect('musculoskeletal_survey.db')
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    
    c = conn.cursor()
    try:
        c.execute('BEGIN TRANSACTION')
        c.execute('''INSERT OR REPLACE INTO survey_data 
                     (session_id, workplace, data, created_at, updated_at) 
                     VALUES (?, ?, ?, datetime('now'), datetime('now'))''', 
                     (session_id, workplace, json.dumps(data, ensure_ascii=False)))
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

# 데이터 불러오기 함수
def load_from_db(session_id):
    conn = sqlite3.connect('musculoskeletal_survey.db')
    c = conn.cursor()
    c.execute('SELECT data FROM survey_data WHERE session_id = ?', (session_id,))
    result = c.fetchone()
    conn.close()
    if result:
        return json.loads(result[0])
    return None

# 작업현장별 세션 관리
if "workplace" not in st.session_state:
    st.session_state["workplace"] = None

if "session_id" not in st.session_state:
    st.session_state["session_id"] = None

# 개선된 자동 저장 기능 (10초마다)
def auto_save():
    if "last_save_time" not in st.session_state:
        st.session_state["last_save_time"] = time.time()
    
    current_time = time.time()
    if current_time - st.session_state["last_save_time"] > 10:  # 10초마다 자동 저장
        save_data = {}
        for key, value in st.session_state.items():
            if isinstance(value, pd.DataFrame):
                save_data[key] = value.to_dict('records')
            elif isinstance(value, (str, int, float, bool, list, dict)):
                save_data[key] = value
            elif hasattr(value, 'isoformat'):
                save_data[key] = value.isoformat()
        
        save_to_db(st.session_state["session_id"], save_data, st.session_state.get("workplace"))
        st.session_state["last_save_time"] = current_time
        st.session_state["last_successful_save"] = datetime.now()

# 값 파싱 함수
def parse_value(value, val_type=float):
    """문자열 값을 숫자로 변환"""
    try:
        if isinstance(value, str):
            value = value.strip()
            if value == "":
                return 0
            value = value.replace(",", "")
            return val_type(value)
        return val_type(value) if value else 0
    except:
        return 0

# 세션 상태 초기화
if "checklist_df" not in st.session_state:
    st.session_state["checklist_df"] = pd.DataFrame()

# 데이터베이스 초기화
init_db()

# 작업명 목록을 가져오는 함수
def get_작업명_목록():
    if not st.session_state["checklist_df"].empty:
        return st.session_state["checklist_df"]["작업명"].dropna().unique().tolist()
    return []

# 단위작업명 목록을 가져오는 함수
def get_단위작업명_목록(작업명=None):
    if not st.session_state["checklist_df"].empty:
        df = st.session_state["checklist_df"]
        if 작업명:
            df = df[df["작업명"] == 작업명]
        return df["단위작업명"].dropna().unique().tolist()
    return []

# 부담작업 설명 매핑 (전역 변수)
부담작업_설명 = {
    "1호": "키보드/마우스 4시간 이상",
    "2호": "같은 동작 2시간 이상 반복",
    "3호": "팔 위/옆으로 2시간 이상",
    "4호": "목/허리 구부림 2시간 이상",
    "5호": "쪼그림/무릎굽힘 2시간 이상",
    "6호": "손가락 집기 2시간 이상",
    "7호": "한손 4.5kg 들기 2시간 이상",
    "8호": "25kg 이상 10회/일",
    "9호": "10kg 이상 25회/일",
    "10호": "4.5kg 이상 분당 2회",
    "11호": "손/무릎 충격 시간당 10회",
    "12호": "정적자세/진동/밀당기기"
}

# 사이드바에 데이터 관리 기능
with st.sidebar:
    st.title("📁 데이터 관리")
    
    # 작업현장 선택/입력
    st.markdown("### 🏭 작업현장 선택")
    작업현장_옵션 = ["현장 선택...", "A사업장", "B사업장", "C사업장", "신규 현장 추가"]
    선택된_현장 = st.selectbox("작업현장", 작업현장_옵션)
    
    if 선택된_현장 == "신규 현장 추가":
        새현장명 = st.text_input("새 현장명 입력")
        if 새현장명:
            st.session_state["workplace"] = 새현장명
            st.session_state["session_id"] = f"{새현장명}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.urandom(4).hex()}"
    elif 선택된_현장 != "현장 선택...":
        st.session_state["workplace"] = 선택된_현장
        if not st.session_state.get("session_id") or 선택된_현장 not in st.session_state.get("session_id", ""):
            st.session_state["session_id"] = f"{선택된_현장}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.urandom(4).hex()}"
    
    # 세션 정보 표시
    if st.session_state.get("session_id"):
        st.info(f"🔐 세션 ID: {st.session_state['session_id']}")
    
    # 자동 저장 상태
    if "last_successful_save" in st.session_state:
        last_save = st.session_state["last_successful_save"]
        st.success(f"✅ 마지막 자동저장: {last_save.strftime('%H:%M:%S')}")
    
    # 수동 저장 버튼
    if st.button("💾 수동 저장", use_container_width=True):
        try:
            save_data = {}
            for key, value in st.session_state.items():
                if isinstance(value, pd.DataFrame):
                    save_data[key] = value.to_dict('records')
                elif isinstance(value, (str, int, float, bool, list, dict)):
                    save_data[key] = value
                elif hasattr(value, 'isoformat'):
                    save_data[key] = value.isoformat()
            
            save_to_db(st.session_state["session_id"], save_data, st.session_state.get("workplace"))
            st.success("✅ 데이터가 서버에 저장되었습니다!")
        except Exception as e:
            st.error(f"저장 중 오류 발생: {str(e)}")
    
    # 이전 세션 불러오기
    st.markdown("---")
    prev_session_id = st.text_input("이전 세션 ID 입력")
    if st.button("📤 이전 세션 불러오기", use_container_width=True):
        if prev_session_id:
            loaded_data = load_from_db(prev_session_id)
            if loaded_data:
                for key, value in loaded_data.items():
                    if isinstance(value, list) and len(value) > 0 and isinstance(value[0], dict):
                        st.session_state[key] = pd.DataFrame(value)
                    else:
                        st.session_state[key] = value
                st.success("✅ 이전 세션 데이터를 불러왔습니다!")
                st.rerun()
            else:
                st.error("해당 세션 ID의 데이터를 찾을 수 없습니다.")
    
    # JSON 파일로 내보내기/가져오기
    st.markdown("---")
    st.subheader("📄 파일로 내보내기/가져오기")
    
    # 내보내기
    if st.button("📥 JSON 파일로 내보내기", use_container_width=True):
        try:
            save_data = {}
            for key, value in st.session_state.items():
                if isinstance(value, pd.DataFrame):
                    save_data[key] = value.to_dict('records')
                elif isinstance(value, (str, int, float, bool, list, dict)):
                    save_data[key] = value
                elif hasattr(value, 'isoformat'):
                    save_data[key] = value.isoformat()
            
            json_str = json.dumps(save_data, ensure_ascii=False, indent=2)
            
            st.download_button(
                label="📥 다운로드",
                data=json_str,
                file_name=f"근골격계조사_{st.session_state.get('workplace', '')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json"
            )
        except Exception as e:
            st.error(f"내보내기 중 오류 발생: {str(e)}")
    
    # 가져오기
    uploaded_file = st.file_uploader("📂 JSON 파일 불러오기", type=['json'])
    if uploaded_file is not None:
        if st.button("📤 데이터 가져오기", use_container_width=True):
            try:
                save_data = json.load(uploaded_file)
                for key, value in save_data.items():
                    if isinstance(value, list) and len(value) > 0 and isinstance(value[0], dict):
                        st.session_state[key] = pd.DataFrame(value)
                    else:
                        st.session_state[key] = value
                st.success("✅ 데이터를 성공적으로 가져왔습니다!")
                st.rerun()
            except Exception as e:
                st.error(f"가져오기 중 오류 발생: {str(e)}")
    
    # 성능 최적화 옵션
    st.markdown("---")
    st.subheader("⚡ 성능 최적화")
    if st.checkbox("대용량 데이터 모드", help="체크리스트가 많을 때 사용하세요"):
        st.session_state["large_data_mode"] = True
    else:
        st.session_state["large_data_mode"] = False
    
    # 부담작업 참고 정보
    with st.expander("📖 부담작업 빠른 참조"):
        st.markdown("""
        **반복동작 관련**
        - 1호: 키보드/마우스 4시간↑
        - 2호: 같은동작 2시간↑ 반복
        - 6호: 손가락집기 2시간↑
        - 7호: 한손 4.5kg 2시간↑
        - 10호: 4.5kg 분당2회↑
        
        **부자연스러운 자세**
        - 3호: 팔 위/옆 2시간↑
        - 4호: 목/허리굽힘 2시간↑
        - 5호: 쪼그림/무릎 2시간↑
        
        **과도한 힘**
        - 8호: 25kg 10회/일↑
        - 9호: 10kg 25회/일↑
        
        **기타**
        - 11호: 손/무릎충격 시간당10회↑
        - 12호: 정적자세/진동/밀당기기
        """)

# 페이지 로드 시 데이터 자동 복구
if "data_loaded" not in st.session_state and st.session_state.get("session_id"):
    saved_data = load_from_db(st.session_state["session_id"])
    if saved_data:
        for key, value in saved_data.items():
            if isinstance(value, list) and len(value) > 0 and isinstance(value[0], dict):
                st.session_state[key] = pd.DataFrame(value)
            else:
                st.session_state[key] = value
        st.session_state["data_loaded"] = True

# 자동 저장 실행
if st.session_state.get("session_id"):
    auto_save()

# 작업현장 선택 확인
if not st.session_state.get("workplace"):
    st.warning("⚠️ 먼저 사이드바에서 작업현장을 선택하거나 입력해주세요!")
    st.stop()

# 메인 화면 시작
st.title(f"근골격계 유해요인조사 - {st.session_state.get('workplace', '')}")

# 탭 정의
tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트",
    "유해요인조사표",
    "작업조건조사",
    "정밀조사",
    "증상조사 분석",
    "작업환경개선계획서"
])

# 1. 사업장개요 탭
with tabs[0]:
    st.title("사업장 개요")
    사업장명 = st.text_input("사업장명", key="사업장명", value=st.session_state.get("workplace", ""))
    소재지 = st.text_input("소재지", key="소재지")
    업종 = st.text_input("업종", key="업종")
    col1, col2 = st.columns(2)
    with col1:
        예비조사 = st.date_input("예비조사일", key="예비조사")
        수행기관 = st.text_input("수행기관", key="수행기관")
    with col2:
        본조사 = st.date_input("본조사일", key="본조사")
        성명 = st.text_input("성명", key="성명")

# 2. 근골격계 부담작업 체크리스트 탭
with tabs[1]:
    st.subheader("근골격계 부담작업 체크리스트")
    
    # 엑셀 파일 업로드 기능 추가
    with st.expander("📤 엑셀 파일 업로드"):
        st.info("""
        📌 엑셀 파일 양식:
        - 첫 번째 열: 작업명
        - 두 번째 열: 단위작업명
        - 3~13번째 열: 1호~11호 (O(해당), △(잠재위험), X(미해당) 중 입력)
        """)
        
        uploaded_excel = st.file_uploader("엑셀 파일 선택", type=['xlsx', 'xls'])
        
        if uploaded_excel is not None:
            try:
                # 엑셀 파일 읽기
                df_excel = pd.read_excel(uploaded_excel)
                
                # 컬럼명 확인 및 조정
                expected_columns = ["작업명", "단위작업명"] + [f"{i}호" for i in range(1, 12)]
                
                # 컬럼 개수가 맞는지 확인
                if len(df_excel.columns) >= 13:
                    # 컬럼명 재설정
                    df_excel.columns = expected_columns[:len(df_excel.columns)]
                    
                    # 값 검증 (O(해당), △(잠재위험), X(미해당)만 허용)
                    valid_values = ["O(해당)", "△(잠재위험)", "X(미해당)"]
                    
                    # 3번째 열부터 13번째 열까지 검증
                    for col in expected_columns[2:]:
                        if col in df_excel.columns:
                            # 유효하지 않은 값은 X(미해당)으로 변경
                            df_excel[col] = df_excel[col].apply(
                                lambda x: x if x in valid_values else "X(미해당)"
                            )
                    
                    if st.button("✅ 데이터 적용하기"):
                        st.session_state["checklist_df"] = df_excel
                        
                        # 즉시 데이터베이스에 저장
                        save_data = {}
                        for key, value in st.session_state.items():
                            if isinstance(value, pd.DataFrame):
                                save_data[key] = value.to_dict('records')
                            elif isinstance(value, (str, int, float, bool, list, dict)):
                                save_data[key] = value
                        
                        save_to_db(st.session_state["session_id"], save_data, st.session_state.get("workplace"))
                        st.session_state["last_save_time"] = time.time()
                        st.session_state["last_successful_save"] = datetime.now()
                        
                        st.success("✅ 엑셀 데이터를 성공적으로 불러오고 저장했습니다!")
                        st.rerun()
                    
                    # 미리보기
                    st.markdown("#### 📋 데이터 미리보기")
                    if st.session_state.get("large_data_mode", False):
                        st.dataframe(df_excel.head(20))
                        st.info(f"전체 {len(df_excel)}개 행 중 상위 20개만 표시됩니다.")
                    else:
                        st.dataframe(df_excel)
                    
                else:
                    st.error("⚠️ 엑셀 파일의 컬럼이 13개 이상이어야 합니다. (작업명, 단위작업명, 1호~11호)")
                    
            except Exception as e:
                st.error(f"❌ 파일 읽기 오류: {str(e)}")
    
    # 샘플 엑셀 파일 다운로드
    with st.expander("📥 샘플 엑셀 파일 다운로드"):
        # 샘플 데이터 생성
        sample_data = pd.DataFrame({
            "작업명": ["조립작업", "조립작업", "포장작업", "포장작업", "운반작업"],
            "단위작업명": ["부품조립", "나사체결", "제품포장", "박스적재", "대차운반"],
            "1호": ["O(해당)", "X(미해당)", "X(미해당)", "O(해당)", "X(미해당)"],
            "2호": ["X(미해당)", "O(해당)", "X(미해당)", "X(미해당)", "O(해당)"],
            "3호": ["△(잠재위험)", "X(미해당)", "O(해당)", "X(미해당)", "X(미해당)"],
            "4호": ["X(미해당)", "X(미해당)", "X(미해당)", "△(잠재위험)", "X(미해당)"],
            "5호": ["X(미해당)", "△(잠재위험)", "X(미해당)", "X(미해당)", "O(해당)"],
            "6호": ["X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)"],
            "7호": ["X(미해당)", "X(미해당)", "△(잠재위험)", "X(미해당)", "X(미해당)"],
            "8호": ["X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)"],
            "9호": ["X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)"],
            "10호": ["X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)", "X(미해당)"],
            "11호": ["O(해당)", "X(미해당)", "X(미해당)", "O(해당)", "△(잠재위험)"]
        })
        
        # 엑셀 파일로 변환
        sample_output = BytesIO()
        with pd.ExcelWriter(sample_output, engine='openpyxl') as writer:
            sample_data.to_excel(writer, sheet_name='체크리스트', index=False)
        
        sample_output.seek(0)
        
        st.download_button(
            label="📥 샘플 엑셀 다운로드",
            data=sample_output,
            file_name="체크리스트_샘플.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("##### 샘플 데이터 구조:")
        st.dataframe(sample_data)
    
    st.markdown("---")
    
    # 기존 데이터 편집기
    columns = [
        "작업명", "단위작업명"
    ] + [f"{i}호" for i in range(1, 12)]
    
    # 세션 상태에 저장된 데이터가 있으면 사용, 없으면 빈 데이터
    if not st.session_state["checklist_df"].empty:
        data = st.session_state["checklist_df"]
    else:
        data = pd.DataFrame(
            columns=columns,
            data=[["", ""] + ["X(미해당)"]*11 for _ in range(5)]
        )

    ho_options = [
        "O(해당)",
        "△(잠재위험)",
        "X(미해당)"
    ]
    column_config = {
        f"{i}호": st.column_config.SelectboxColumn(
            f"{i}호", options=ho_options, required=True
        ) for i in range(1, 12)
    }
    column_config["작업명"] = st.column_config.TextColumn("작업명")
    column_config["단위작업명"] = st.column_config.TextColumn("단위작업명")

    # 대용량 데이터 모드에서는 페이지네이션 사용
    if st.session_state.get("large_data_mode", False) and len(data) > 50:
        st.warning("대용량 데이터 모드가 활성화되었습니다.")
        
        # 페이지네이션 개선
        page_size = st.selectbox("페이지당 행 수", [25, 50, 100, 200], index=1)
        total_pages = (len(data) - 1) // page_size + 1
        
        # 페이지 네비게이션
        col1, col2, col3 = st.columns([2, 3, 2])
        with col1:
            if st.button("◀ 이전", disabled=(st.session_state.get('current_page', 1) <= 1)):
                st.session_state['current_page'] = st.session_state.get('current_page', 1) - 1
                st.rerun()
        
        with col2:
            page = st.selectbox(
                "페이지", 
                range(1, total_pages + 1), 
                index=st.session_state.get('current_page', 1) - 1,
                format_func=lambda x: f"{x}/{total_pages}"
            )
            st.session_state['current_page'] = page
        
        with col3:
            if st.button("다음 ▶", disabled=(st.session_state.get('current_page', 1) >= total_pages)):
                st.session_state['current_page'] = st.session_state.get('current_page', 1) + 1
                st.rerun()
        
        start_idx = (page - 1) * page_size
        end_idx = min(start_idx + page_size, len(data))
        
        # 현재 페이지 데이터만 표시
        page_data = data.iloc[start_idx:end_idx].copy()
        
        edited_df = st.data_editor(
            page_data,
            use_container_width=True,
            hide_index=True,
            column_config=column_config,
            key=f"page_editor_{page}"
        )
        
        # 편집된 데이터 병합
        data.iloc[start_idx:end_idx] = edited_df
        st.session_state["checklist_df"] = data
        
        # 전체 데이터 요약 표시
        st.info(f"📊 전체 {len(data)}개 행 중 {start_idx+1}-{end_idx}번째 표시 중")
        
        # 빠른 검색 기능
        search_col1, search_col2 = st.columns([1, 3])
        with search_col1:
            search_field = st.selectbox("검색 필드", ["작업명", "단위작업명"])
        with search_col2:
            search_term = st.text_input("검색어", key="checklist_search")
        
        if search_term:
            filtered_data = data[data[search_field].str.contains(search_term, case=False, na=False)]
            st.write(f"🔍 '{search_term}' 검색 결과: {len(filtered_data)}개")
            if len(filtered_data) > 0:
                st.dataframe(filtered_data.head(10))
    else:
        edited_df = st.data_editor(
            data,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config=column_config
        )
        st.session_state["checklist_df"] = edited_df
    
    # 현재 등록된 작업명 표시
    작업명_목록 = get_작업명_목록()
    if 작업명_목록:
        st.info(f"📋 현재 등록된 작업: {', '.join(작업명_목록)}")

# 3. 유해요인조사표 탭
with tabs[2]:
    st.title("유해요인조사표")
    
    # 작업명 목록 가져오기
    작업명_목록 = get_작업명_목록()
    
    if not 작업명_목록:
        st.warning("⚠️ 먼저 '근골격계 부담작업 체크리스트' 탭에서 작업명을 입력해주세요.")
    else:
        # 작업명별로 유해요인조사표 작성
        selected_작업명_유해 = st.selectbox(
            "작업명 선택",
            작업명_목록,
            key="작업명_선택_유해요인"
        )
        
        st.info(f"📋 선택된 작업: {selected_작업명_유해}")
        
        # 해당 작업의 단위작업명 가져오기
        단위작업명_목록 = get_단위작업명_목록(selected_작업명_유해)
        
        with st.expander(f"📌 {selected_작업명_유해} - 유해요인조사표", expanded=True):
            st.markdown("#### 가. 조사개요")
            col1, col2 = st.columns(2)
            with col1:
                조사일시 = st.text_input("조사일시", key=f"조사일시_{selected_작업명_유해}")
                부서명 = st.text_input("부서명", key=f"부서명_{selected_작업명_유해}")
            with col2:
                조사자 = st.text_input("조사자", key=f"조사자_{selected_작업명_유해}")
                작업공정명 = st.text_input("작업공정명", value=selected_작업명_유해, key=f"작업공정명_{selected_작업명_유해}")
            작업명_유해 = st.text_input("작업명", value=selected_작업명_유해, key=f"작업명_{selected_작업명_유해}")
            
            # 단위작업명 표시
            if 단위작업명_목록:
                st.markdown("##### 단위작업명 목록")
                st.write(", ".join(단위작업명_목록))

            st.markdown("#### 나. 작업장 상황조사")

            def 상황조사행(항목명, 작업명):
                cols = st.columns([2, 5, 3])
                with cols[0]:
                    st.markdown(f"<div style='text-align:center; font-weight:bold; padding-top:0.7em;'>{항목명}</div>", unsafe_allow_html=True)
                with cols[1]:
                    상태 = st.radio(
                        label="",
                        options=["변화없음", "감소", "증가", "기타"],
                        key=f"{항목명}_상태_{작업명}",
                        horizontal=True,
                        label_visibility="collapsed"
                    )
                with cols[2]:
                    if 상태 == "감소":
                        st.text_input("감소 - 언제부터", key=f"{항목명}_감소_시작_{작업명}", placeholder="언제부터", label_visibility="collapsed")
                    elif 상태 == "증가":
                        st.text_input("증가 - 언제부터", key=f"{항목명}_증가_시작_{작업명}", placeholder="언제부터", label_visibility="collapsed")
                    elif 상태 == "기타":
                        st.text_input("기타 - 내용", key=f"{항목명}_기타_내용_{작업명}", placeholder="내용", label_visibility="collapsed")
                    else:
                        st.markdown("&nbsp;", unsafe_allow_html=True)

            for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
                상황조사행(항목, selected_작업명_유해)
                st.markdown("<hr style='margin:0.5em 0;'>", unsafe_allow_html=True)
            
            st.markdown("---")

# 작업부하와 작업빈도에서 숫자 추출하는 함수
def extract_number(value):
    if value and "(" in value and ")" in value:
        return int(value.split("(")[1].split(")")[0])
    return 0

# 총점 계산 함수
def calculate_total_score(row):
    부하값 = extract_number(row["작업부하(A)"])
    빈도값 = extract_number(row["작업빈도(B)"])
    return 부하값 * 빈도값

# 4. 작업조건조사 탭
with tabs[3]:
    st.title("작업조건조사")
    
    # 체크리스트에서 작업명 목록 가져오기
    작업명_목록 = get_작업명_목록()
    
    if not 작업명_목록:
        st.warning("⚠️ 먼저 '근골격계 부담작업 체크리스트' 탭에서 작업명을 입력해주세요.")
    else:
        # 작업명 선택
        selected_작업명 = st.selectbox(
            "작업명 선택",
            작업명_목록,
            key="작업명_선택"
        )
        
        st.info(f"📋 총 {len(작업명_목록)}개의 작업이 있습니다. 각 작업별로 1,2,3단계를 작성하세요.")
        
        # 선택된 작업에 대한 1,2,3단계
        with st.container():
            # 1단계: 유해요인 기본조사
            st.subheader(f"1단계: 유해요인 기본조사 - [{selected_작업명}]")
            col1, col2 = st.columns(2)
            with col1:
                작업공정 = st.text_input("작업공정", value=selected_작업명, key=f"1단계_작업공정_{selected_작업명}")
            with col2:
                작업내용 = st.text_input("작업내용", key=f"1단계_작업내용_{selected_작업명}")
            
            st.markdown("---")
            
            # 2단계: 작업별 작업부하 및 작업빈도
            st.subheader(f"2단계: 작업별 작업부하 및 작업빈도 - [{selected_작업명}]")
            
            # 선택된 작업명에 해당하는 체크리스트 데이터 가져오기
            checklist_data = []
            if not st.session_state["checklist_df"].empty:
                작업_체크리스트 = st.session_state["checklist_df"][
                    st.session_state["checklist_df"]["작업명"] == selected_작업명
                ]
                
                for idx, row in 작업_체크리스트.iterrows():
                    if row["단위작업명"]:
                        부담작업호 = []
                        for i in range(1, 12):
                            if row[f"{i}호"] == "O(해당)":
                                부담작업호.append(f"{i}호")
                            elif row[f"{i}호"] == "△(잠재위험)":
                                부담작업호.append(f"{i}호(잠재)")
                        
                        checklist_data.append({
                            "단위작업명": row["단위작업명"],
                            "부담작업(호)": ", ".join(부담작업호) if 부담작업호 else "미해당",
                            "작업부하(A)": "",
                            "작업빈도(B)": "",
                            "총점": 0
                        })
            
            # 데이터프레임 생성
            if checklist_data:
                data = pd.DataFrame(checklist_data)
            else:
                data = pd.DataFrame({
                    "단위작업명": ["" for _ in range(3)],
                    "부담작업(호)": ["" for _ in range(3)],
                    "작업부하(A)": ["" for _ in range(3)],
                    "작업빈도(B)": ["" for _ in range(3)],
                    "총점": [0 for _ in range(3)],
                })

            부하옵션 = [
                "",
                "매우쉬움(1)", 
                "쉬움(2)", 
                "약간 힘듦(3)", 
                "힘듦(4)", 
                "매우 힘듦(5)"
            ]
            빈도옵션 = [
                "",
                "3개월마다(1)", 
                "가끔(2)", 
                "자주(3)", 
                "계속(4)", 
                "초과근무(5)"
            ]

            column_config = {
                "작업부하(A)": st.column_config.SelectboxColumn("작업부하(A)", options=부하옵션, required=False),
                "작업빈도(B)": st.column_config.SelectboxColumn("작업빈도(B)", options=빈도옵션, required=False),
                "단위작업명": st.column_config.TextColumn("단위작업명"),
                "부담작업(호)": st.column_config.TextColumn("부담작업(호)"),
                "총점": st.column_config.TextColumn("총점(자동계산)", disabled=True),
            }

            # 데이터 편집
            edited_df = st.data_editor(
                data,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config=column_config,
                key=f"작업조건_data_editor_{selected_작업명}"
            )
            
            # 편집된 데이터를 세션 상태에 저장
            st.session_state[f"작업조건_data_{selected_작업명}"] = edited_df
            
            # 총점 자동 계산 후 다시 표시
            if not edited_df.empty:
                display_df = edited_df.copy()
                for idx in range(len(display_df)):
                    display_df.at[idx, "총점"] = calculate_total_score(display_df.iloc[idx])
                
                st.markdown("##### 계산 결과")
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "단위작업명": st.column_config.TextColumn("단위작업명"),
                        "부담작업(호)": st.column_config.TextColumn("부담작업(호)"),
                        "작업부하(A)": st.column_config.TextColumn("작업부하(A)"),
                        "작업빈도(B)": st.column_config.TextColumn("작업빈도(B)"),
                        "총점": st.column_config.NumberColumn("총점(자동계산)", format="%d"),
                    }
                )
                
                st.info("💡 총점은 작업부하(A) × 작업빈도(B)로 자동 계산됩니다.")
            
            # 3단계: 유해요인평가
            st.markdown("---")
            st.subheader(f"3단계: 유해요인평가 - [{selected_작업명}]")
            
            # 작업명과 근로자수 입력
            col1, col2 = st.columns(2)
            with col1:
                평가_작업명 = st.text_input("작업명", value=selected_작업명, key=f"3단계_작업명_{selected_작업명}")
            with col2:
                평가_근로자수 = st.text_input("근로자수", key=f"3단계_근로자수_{selected_작업명}")
            
            # 사진 업로드 및 설명 입력
            st.markdown("#### 작업 사진 및 설명")
            
            # 사진 개수 선택
            num_photos = st.number_input("사진 개수", min_value=1, max_value=10, value=3, key=f"사진개수_{selected_작업명}")
            
            # 각 사진별로 업로드와 설명 입력
            for i in range(num_photos):
                st.markdown(f"##### 사진 {i+1}")
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    uploaded_file = st.file_uploader(
                        f"사진 {i+1} 업로드",
                        type=['png', 'jpg', 'jpeg'],
                        key=f"사진_{i+1}_업로드_{selected_작업명}"
                    )
                    if uploaded_file:
                        st.image(uploaded_file, caption=f"사진 {i+1}", use_column_width=True)
                
                with col2:
                    photo_description = st.text_area(
                        f"사진 {i+1} 설명",
                        height=150,
                        key=f"사진_{i+1}_설명_{selected_작업명}",
                        placeholder="이 사진에 대한 설명을 입력하세요..."
                    )
                
                st.markdown("---")
            
            # 작업별로 관련된 유해요인에 대한 원인분석 (개선된 버전)
            st.markdown("---")
            st.subheader(f"작업별로 관련된 유해요인에 대한 원인분석 - [{selected_작업명}]")
            
            # 2단계에서 입력한 데이터와 체크리스트 정보 가져오기
            부담작업_정보 = []
            부담작업_힌트 = {}  # 단위작업명별 부담작업 정보 저장
            
            if 'display_df' in locals() and not display_df.empty:
                for idx, row in display_df.iterrows():
                    if row["단위작업명"] and row["부담작업(호)"] and row["부담작업(호)"] != "미해당":
                        부담작업_정보.append({
                            "단위작업명": row["단위작업명"],
                            "부담작업호": row["부담작업(호)"]
                        })
                        부담작업_힌트[row["단위작업명"]] = row["부담작업(호)"]
            
            # 원인분석 항목 초기화
            원인분석_key = f"원인분석_항목_{selected_작업명}"
            if 원인분석_key not in st.session_state:
                st.session_state[원인분석_key] = []
                # 부담작업 정보를 기반으로 초기 항목 생성 (부담작업이 있는 개수만큼)
                for info in 부담작업_정보:
                    st.session_state[원인분석_key].append({
                        "단위작업명": info["단위작업명"],
                        "부담작업호": info["부담작업호"],
                        "유형": "",
                        "부담작업": "",
                        "비고": ""
                    })
            
            # 추가/삭제 버튼
            col1, col2, col3 = st.columns([6, 1, 1])
            with col2:
                if st.button("➕ 추가", key=f"원인분석_추가_{selected_작업명}", use_container_width=True):
                    st.session_state[원인분석_key].append({
                        "단위작업명": "",
                        "부담작업호": "",
                        "유형": "",
                        "부담작업": "",
                        "비고": ""
                    })
                    st.rerun()
            with col3:
                if st.button("➖ 삭제", key=f"원인분석_삭제_{selected_작업명}", use_container_width=True):
                    if len(st.session_state[원인분석_key]) > 0:
                        st.session_state[원인분석_key].pop()
                        st.rerun()
            
            # 유형별 관련 부담작업 매핑
            유형별_부담작업 = {
                "반복동작": ["1호", "2호", "6호", "7호", "10호"],
                "부자연스러운 자세": ["3호", "4호", "5호"],
                "과도한 힘": ["8호", "9호"],
                "접촉스트레스 또는 기타(진동, 밀고 당기기 등)": ["11호", "12호"]
            }
            
            # 각 유해요인 항목 처리
            hazard_entries_to_process = st.session_state[원인분석_key]
            
            for k, hazard_entry in enumerate(hazard_entries_to_process):
                st.markdown(f"**유해요인 원인분석 항목 {k+1}**")
                
                # 단위작업명 입력 및 부담작업 힌트 표시
                col1, col2, col3 = st.columns([3, 2, 3])
                
                with col1:
                    hazard_entry["단위작업명"] = st.text_input(
                        "단위작업명", 
                        value=hazard_entry.get("단위작업명", ""), 
                        key=f"원인분석_단위작업명_{k}_{selected_작업명}"
                    )
                
                with col2:
                    # 해당 단위작업의 부담작업 정보를 힌트로 표시
                    if hazard_entry["단위작업명"] in 부담작업_힌트:
                        부담작업_리스트 = 부담작업_힌트[hazard_entry["단위작업명"]].split(", ")
                        힌트_텍스트= []
                        
                        for 항목 in 부담작업_리스트:
                            호수 = 항목.replace("(잠재)", "").strip()
                            if 호수 in 부담작업_설명:
                                if "(잠재)" in 항목:
                                    힌트_텍스트.append(f"🟡 {호수}: {부담작업_설명[호수]}")
                                else:
                                    힌트_텍스트.append(f"🔴 {호수}: {부담작업_설명[호수]}")
                        
                        if 힌트_텍스트:
                            st.info("💡 부담작업 힌트:\n" + "\n".join(힌트_텍스트))
                    else:
                        st.empty()  # 빈 공간 유지
                
                with col3:
                    hazard_entry["비고"] = st.text_input(
                        "비고", 
                        value=hazard_entry.get("비고", ""), 
                        key=f"원인분석_비고_{k}_{selected_작업명}"
                    )
                
                # 유해요인 유형 선택
                hazard_type_options = ["", "반복동작", "부자연스러운 자세", "과도한 힘", "접촉스트레스 또는 기타(진동, 밀고 당기기 등)"]
                selected_hazard_type_index = hazard_type_options.index(hazard_entry.get("유형", "")) if hazard_entry.get("유형", "") in hazard_type_options else 0
                
                hazard_entry["유형"] = st.selectbox(
                    f"[{k+1}] 유해요인 유형 선택", 
                    hazard_type_options, 
                    index=selected_hazard_type_index, 
                    key=f"hazard_type_{k}_{selected_작업명}",
                    help="선택한 단위작업의 부담작업 유형에 맞는 항목을 선택하세요"
                )

                if hazard_entry["유형"] == "반복동작":
                    burden_task_options = [
                        "",
                        "(1호)하루에 4시간 이상 집중적으로 자료입력 등을 위해 키보드 또는 마우스를 조작하는 작업",
                        "(2호)하루에 총 2시간 이상 목, 어깨, 팔꿈치, 손목 또는 손을 사용하여 같은 동작을 반복하는 작업",
                        "(6호)하루에 총 2시간 이상 지지되지 않은 상태에서 1kg 이상의 물건을 한손의 손가락으로 집어 옮기거나, 2kg 이상에 상응하는 힘을 가하여 한손의 손가락으로 물건을 쥐는 작업",
                        "(7호)하루에 총 2시간 이상 지지되지 않은 상태에서 4.5kg 이상의 물건을 한 손으로 들거나 동일한 힘으로 쥐는 작업",
                        "(10호)하루에 총 2시간 이상, 분당 2회 이상 4.5kg 이상의 물체를 드는 작업",
                        "(1호)하루에 4시간 이상 집중적으로 자료입력 등을 위해 키보드 또는 마우스를 조작하는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                        "(2호)하루에 총 2시간 이상 목, 어깨, 팔꿈치, 손목 또는 손을 사용하여 같은 동작을 반복하는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                        "(6호)하루에 총 2시간 이상 지지되지 않은 상태에서 1kg 이상의 물건을 한손의 손가락으로 집어 옮기거나, 2kg 이상에 상응하는 힘을 가하여 한손의 손가락으로 물건을 쥐는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                        "(7호)하루에 총 2시간 이상 지지되지 않은 상태에서 4.5kg 이상의 물건을 한 손으로 들거나 동일한 힘으로 쥐는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)",
                        "(10호)하루에 총 2시간 이상, 분당 2회 이상 4.5kg 이상의 물체를 드는 작업+(12호)정적자세(장시간 서서 작업, 또는 장시간 앉아서 작업)"
                    ]
                    selected_burden_task_index = burden_task_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_task_options else 0
                    hazard_entry["부담작업"] = st.selectbox(f"[{k+1}] 부담작업", burden_task_options, index=selected_burden_task_index, key=f"burden_task_반복_{k}_{selected_작업명}")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        hazard_entry["수공구 종류"] = st.text_input(f"[{k+1}] 수공구 종류", value=hazard_entry.get("수공구 종류", ""), key=f"수공구_종류_{k}_{selected_작업명}")
                        hazard_entry["부담부위"] = st.text_input(f"[{k+1}] 부담부위", value=hazard_entry.get("부담부위", ""), key=f"부담부위_{k}_{selected_작업명}")
                    with col2:
                        hazard_entry["수공구 용도"] = st.text_input(f"[{k+1}] 수공구 용도", value=hazard_entry.get("수공구 용도", ""), key=f"수공구_용도_{k}_{selected_작업명}")
                        회당_반복시간_초_회 = st.text_input(f"[{k+1}] 회당 반복시간(초/회)", value=hazard_entry.get("회당 반복시간(초/회)", ""), key=f"반복_회당시간_{k}_{selected_작업명}")
                    with col3:
                        hazard_entry["수공구 무게(kg)"] = st.number_input(f"[{k+1}] 수공구 무게(kg)", value=hazard_entry.get("수공구 무게(kg)", 0.0), key=f"수공구_무게_{k}_{selected_작업명}")
                        작업시간동안_반복횟수_회_일 = st.text_input(f"[{k+1}] 작업시간동안 반복횟수(회/일)", value=hazard_entry.get("작업시간동안 반복횟수(회/일)", ""), key=f"반복_총횟수_{k}_{selected_작업명}")
                    with col4:
                        hazard_entry["수공구 사용시간(분)"] = st.text_input(f"[{k+1}] 수공구 사용시간(분)", value=hazard_entry.get("수공구 사용시간(분)", ""), key=f"수공구_사용시간_{k}_{selected_작업명}")
                        
                        # 총 작업시간(분) 자동 계산
                        calculated_total_work_time = 0.0
                        try:
                            parsed_회당_반복시간 = parse_value(회당_반복시간_초_회, val_type=float)
                            parsed_작업시간동안_반복횟수 = parse_value(작업시간동안_반복횟수_회_일, val_type=float)
                            
                            if parsed_회당_반복시간 > 0 and parsed_작업시간동안_반복횟수 > 0:
                                calculated_total_work_time = (parsed_회당_반복시간 * parsed_작업시간동안_반복횟수) / 60
                        except Exception:
                            pass
                        
                        hazard_entry["총 작업시간(분)"] = st.text_input(
                            f"[{k+1}] 총 작업시간(분) (자동계산)",
                            value=f"{calculated_total_work_time:.2f}" if calculated_total_work_time > 0 else "",
                            key=f"반복_총시간_{k}_{selected_작업명}",
                            disabled=True
                        )
                    
                    # 값 저장
                    hazard_entry["회당 반복시간(초/회)"] = 회당_반복시간_초_회
                    hazard_entry["작업시간동안 반복횟수(회/일)"] = 작업시간동안_반복횟수_회_일

                    # 10호 추가 필드
                    if "(10호)" in hazard_entry["부담작업"]:
                        col1, col2 = st.columns(2)
                        with col1:
                            hazard_entry["물체 무게(kg)_10호"] = st.number_input(f"[{k+1}] (10호)물체 무게(kg)", value=hazard_entry.get("물체 무게(kg)_10호", 0.0), key=f"물체_무게_10호_{k}_{selected_작업명}")
                        with col2:
                            hazard_entry["분당 반복횟수(회/분)_10호"] = st.text_input(f"[{k+1}] (10호)분당 반복횟수(회/분)", value=hazard_entry.get("분당 반복횟수(회/분)_10호", ""), key=f"분당_반복횟수_10호_{k}_{selected_작업명}")

                    # 12호 정적자세 관련 필드
                    if "(12호)정적자세" in hazard_entry["부담작업"]:
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            hazard_entry["작업내용_12호_정적"] = st.text_input(f"[{k+1}] (정지자세)작업내용", value=hazard_entry.get("작업내용_12호_정적", ""), key=f"반복_작업내용_12호_정적_{k}_{selected_작업명}")
                        with col2:
                            hazard_entry["작업시간(분)_12호_정적"] = st.number_input(f"[{k+1}] (정지자세)작업시간(분)", value=hazard_entry.get("작업시간(분)_12호_정적", 0), key=f"반복_작업시간_12호_정적_{k}_{selected_작업명}")
                        with col3:
                            hazard_entry["휴식시간(분)_12호_정적"] = st.number_input(f"[{k+1}] (정지자세)휴식시간(분)", value=hazard_entry.get("휴식시간(분)_12호_정적", 0), key=f"반복_휴식시간_12호_정적_{k}_{selected_작업명}")
                        with col4:
                            hazard_entry["인체부담부위_12호_정적"] = st.text_input(f"[{k+1}] (정지자세)인체부담부위", value=hazard_entry.get("인체부담부위_12호_정적", ""), key=f"반복_인체부담부위_12호_정적_{k}_{selected_작업명}")

                elif hazard_entry["유형"] == "부자연스러운 자세":
                    burden_pose_options = [
                        "",
                        "(3호)하루에 총 2시간 이상 머리 위에 손이 있거나, 팔꿈치가 어깨위에 있거나, 팔꿈치를 몸통으로부터 들거나, 팔꿈치를 몸통뒤쪽에 위치하도록 하는 상태에서 이루어지는 작업",
                        "(4호)지지되지 않은 상태이거나 임의로 자세를 바꿀 수 없는 조건에서, 하루에 총 2시간 이상 목이나 허리를 구부리거나 트는 상태에서 이루어지는 작업",
                        "(5호)하루에 총 2시간 이상 쪼그리고 앉거나 무릎을 굽힌 자세에서 이루어지는 작업"
                    ]
                    selected_burden_pose_index = burden_pose_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_pose_options else 0
                    hazard_entry["부담작업"] = st.selectbox(f"[{k+1}] 부담작업", burden_pose_options, index=selected_burden_pose_index, key=f"burden_pose_{k}_{selected_작업명}")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        hazard_entry["회당 반복시간(초/회)"] = st.text_input(f"[{k+1}] 회당 반복시간(초/회)", value=hazard_entry.get("회당 반복시간(초/회)", ""), key=f"자세_회당시간_{k}_{selected_작업명}")
                    with col2:
                        hazard_entry["작업시간동안 반복횟수(회/일)"] = st.text_input(f"[{k+1}] 작업시간동안 반복횟수(회/일)", value=hazard_entry.get("작업시간동안 반복횟수(회/일)", ""), key=f"자세_총횟수_{k}_{selected_작업명}")
                    with col3:
                        hazard_entry["총 작업시간(분)"] = st.text_input(f"[{k+1}] 총 작업시간(분)", value=hazard_entry.get("총 작업시간(분)", ""), key=f"자세_총시간_{k}_{selected_작업명}")

                elif hazard_entry["유형"] == "과도한 힘":
                    burden_force_options = [
                        "",
                        "(8호)하루에 10회 이상 25kg 이상의 물체를 드는 작업",
                        "(9호)하루에 25회 이상 10kg 이상의 물체를 무릎 아래에서 들거나, 어깨 위에서 들거나, 팔을 뻗은 상태에서 드는 작업",
                        "(12호)밀기/당기기 작업",
                        "(8호)하루에 10회 이상 25kg 이상의 물체를 드는 작업+(12호)밀기/당기기 작업",
                        "(9호)하루에 25회 이상 10kg 이상의 물체를 무릎 아래에서 들거나, 어깨 위에서 들거나, 팔을 뻗은 상태에서 드는 작업+(12호)밀기/당기기 작업"
                    ]
                    selected_burden_force_index = burden_force_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_force_options else 0
                    hazard_entry["부담작업"] = st.selectbox(f"[{k+1}] 부담작업", burden_force_options, index=selected_burden_force_index, key=f"burden_force_{k}_{selected_작업명}")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        hazard_entry["중량물 명칭"] = st.text_input(f"[{k+1}] 중량물 명칭", value=hazard_entry.get("중량물 명칭", ""), key=f"힘_중량물_명칭_{k}_{selected_작업명}")
                    with col2:
                        hazard_entry["중량물 용도"] = st.text_input(f"[{k+1}] 중량물 용도", value=hazard_entry.get("중량물 용도", ""), key=f"힘_중량물_용도_{k}_{selected_작업명}")
                    
                    # 취급방법
                    취급방법_options = ["", "직접 취급", "크레인 사용"]
                    selected_취급방법_index = 취급방법_options.index(hazard_entry.get("취급방법", "")) if hazard_entry.get("취급방법", "") in 취급방법_options else 0
                    hazard_entry["취급방법"] = st.selectbox(f"[{k+1}] 취급방법", 취급방법_options, index=selected_취급방법_index, key=f"힘_취급방법_{k}_{selected_작업명}")

                    # 중량물 이동방법 (취급방법이 "직접 취급"인 경우만 해당)
                    if hazard_entry["취급방법"] == "직접 취급":
                        이동방법_options = ["", "1인 직접이동", "2인1조 직접이동", "여러명 직접이동", "이동대차(인력이동)", "이동대차(전력이동)", "지게차"]
                        selected_이동방법_index = 이동방법_options.index(hazard_entry.get("중량물 이동방법", "")) if hazard_entry.get("중량물 이동방법", "") in 이동방법_options else 0
                        hazard_entry["중량물 이동방법"] = st.selectbox(f"[{k+1}] 중량물 이동방법", 이동방법_options, index=selected_이동방법_index, key=f"힘_이동방법_{k}_{selected_작업명}")
                        
                        # 이동대차(인력이동) 선택 시 추가 드롭다운
                        if hazard_entry["중량물 이동방법"] == "이동대차(인력이동)":
                            직접_밀당_options = ["", "작업자가 직접 바퀴달린 이동대차를 밀고/당기기", "자동이동대차(AGV)", "기타"]
                            selected_직접_밀당_index = 직접_밀당_options.index(hazard_entry.get("작업자가 직접 밀고/당기기", "")) if hazard_entry.get("작업자가 직접 밀고/당기기", "") in 직접_밀당_options else 0
                            hazard_entry["작업자가 직접 밀고/당기기"] = st.selectbox(f"[{k+1}] 작업자가 직접 밀고/당기기", 직접_밀당_options, index=selected_직접_밀당_index, key=f"힘_직접_밀당_{k}_{selected_작업명}")
                            # '기타' 선택 시 설명 적는 난 추가
                            if hazard_entry["작업자가 직접 밀고/당기기"] == "기타":
                                hazard_entry["기타_밀당_설명"] = st.text_input(f"[{k+1}] 기타 밀기/당기기 설명", value=hazard_entry.get("기타_밀당_설명", ""), key=f"힘_기타_밀당_설명_{k}_{selected_작업명}")

                    # 8호, 9호 관련 필드 (밀기/당기기가 아닌 경우)
                    if "(8호)" in hazard_entry["부담작업"] and "(12호)" not in hazard_entry["부담작업"]:
                        col1, col2 = st.columns(2)
                        with col1:
                            hazard_entry["중량물 무게(kg)"] = st.number_input(f"[{k+1}] 중량물 무게(kg)", value=hazard_entry.get("중량물 무게(kg)", 0.0), key=f"중량물_무게_{k}_{selected_작업명}")
                        with col2:
                            hazard_entry["작업시간동안 작업횟수(회/일)"] = st.text_input(f"[{k+1}] 작업시간동안 작업횟수(회/일)", value=hazard_entry.get("작업시간동안 작업횟수(회/일)", ""), key=f"힘_총횟수_{k}_{selected_작업명}")
                    
                    elif "(9호)" in hazard_entry["부담작업"] and "(12호)" not in hazard_entry["부담작업"]:
                        col1, col2 = st.columns(2)
                        with col1:
                            hazard_entry["중량물 무게(kg)"] = st.number_input(f"[{k+1}] 중량물 무게(kg)", value=hazard_entry.get("중량물 무게(kg)", 0.0), key=f"중량물_무게_{k}_{selected_작업명}")
                        with col2:
                            hazard_entry["작업시간동안 작업횟수(회/일)"] = st.text_input(f"[{k+1}] 작업시간동안 작업횟수(회/일)", value=hazard_entry.get("작업시간동안 작업횟수(회/일)", ""), key=f"힘_총횟수_{k}_{selected_작업명}")
                    
                    # 12호 밀기/당기기 관련 필드
                    if "(12호)밀기/당기기" in hazard_entry["부담작업"]:
                        st.markdown("##### (12호) 밀기/당기기 세부 정보")
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            hazard_entry["대차 무게(kg)_12호"] = st.number_input(f"[{k+1}] 대차 무게(kg)", value=hazard_entry.get("대차 무게(kg)_12호", 0.0), key=f"대차_무게_12호_{k}_{selected_작업명}")
                        with col2:
                            hazard_entry["대차위 제품무게(kg)_12호"] = st.number_input(f"[{k+1}] 대차위 제품무게(kg)", value=hazard_entry.get("대차위 제품무게(kg)_12호", 0.0), key=f"대차위_제품무게_12호_{k}_{selected_작업명}")
                        with col3:
                            hazard_entry["밀고-당기기 빈도(회/일)_12호"] = st.text_input(f"[{k+1}] 밀고-당기기 빈도(회/일)", value=hazard_entry.get("밀고-당기기 빈도(회/일)_12호", ""), key=f"밀고당기기_빈도_12호_{k}_{selected_작업명}")

                elif hazard_entry["유형"] == "접촉스트레스 또는 기타(진동, 밀고 당기기 등)":
                    burden_other_options = [
                        "",
                        "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업",
                        "(12호)진동작업(그라인더, 임팩터 등)"
                    ]
                    selected_burden_other_index = burden_other_options.index(hazard_entry.get("부담작업", "")) if hazard_entry.get("부담작업", "") in burden_other_options else 0
                    hazard_entry["부담작업"] = st.selectbox(f"[{k+1}] 부담작업", burden_other_options, index=selected_burden_other_index, key=f"burden_other_{k}_{selected_작업명}")

                    if hazard_entry["부담작업"] == "(11호)하루에 총 2시간 이상 시간당 10회 이상 손 또는 무릎을 사용하여 반복적으로 충격을 가하는 작업":
                        hazard_entry["작업시간(분)"] = st.text_input(f"[{k+1}] 작업시간(분)", value=hazard_entry.get("작업시간(분)", ""), key=f"기타_작업시간_{k}_{selected_작업명}")

                    if hazard_entry["부담작업"] == "(12호)진동작업(그라인더, 임팩터 등)":
                        st.markdown("##### (12호) 진동작업 세부 정보")
                        col1, col2 = st.columns(2)
                        with col1:
                            hazard_entry["진동수공구명"] = st.text_input(f"[{k+1}] 진동수공구명", value=hazard_entry.get("진동수공구명", ""), key=f"기타_진동수공구명_{k}_{selected_작업명}")
                            hazard_entry["작업시간(분)_진동"] = st.text_input(f"[{k+1}] 작업시간(분)", value=hazard_entry.get("작업시간(분)_진동", ""), key=f"기타_작업시간_진동_{k}_{selected_작업명}")
                            hazard_entry["작업량(회/일)_진동"] = st.text_input(f"[{k+1}] 작업량(회/일)", value=hazard_entry.get("작업량(회/일)_진동", ""), key=f"기타_작업량_진동_{k}_{selected_작업명}")
                        with col2:
                            hazard_entry["진동수공구 용도"] = st.text_input(f"[{k+1}] 진동수공구 용도", value=hazard_entry.get("진동수공구 용도", ""), key=f"기타_진동수공구_용도_{k}_{selected_작업명}")
                            hazard_entry["작업빈도(초/회)_진동"] = st.text_input(f"[{k+1}] 작업빈도(초/회)", value=hazard_entry.get("작업빈도(초/회)_진동", ""), key=f"기타_작업빈도_진동_{k}_{selected_작업명}")
                            
                            지지대_options = ["", "예", "아니오"]
                            selected_지지대_index = 지지대_options.index(hazard_entry.get("수공구사용시 지지대가 있는가?", "")) if hazard_entry.get("수공구사용시 지지대가 있는가?", "") in 지지대_options else 0
                            hazard_entry["수공구사용시 지지대가 있는가?"] = st.selectbox(f"[{k+1}] 수공구사용시 지지대가 있는가?", 지지대_options, index=selected_지지대_index, key=f"기타_지지대_여부_{k}_{selected_작업명}")
                
                st.markdown("---")

# 5. 정밀조사 탭
with tabs[4]:
    st.title("정밀조사")
    
    # 세션 상태 초기화
    if "정밀조사_목록" not in st.session_state:
        st.session_state["정밀조사_목록"] = []
    
    # 정밀조사 추가 버튼
    col1, col2 = st.columns([6, 1])
    with col2:
        if st.button("➕ 정밀조사 추가", use_container_width=True):
            st.session_state["정밀조사_목록"].append(f"정밀조사_{len(st.session_state['정밀조사_목록'])+1}")
            st.rerun()
    
    if not st.session_state["정밀조사_목록"]:
        st.info("📋 정밀조사가 필요한 경우 '정밀조사 추가' 버튼을 클릭하세요.")
    else:
        # 각 정밀조사 표시
        for idx, 조사명 in enumerate(st.session_state["정밀조사_목록"]):
            with st.expander(f"📌 {조사명}", expanded=True):
                # 삭제 버튼
                col1, col2 = st.columns([10, 1])
                with col2:
                    if st.button("❌", key=f"삭제_{조사명}"):
                        st.session_state["정밀조사_목록"].remove(조사명)
                        st.rerun()
                
                # 정밀조사표
                st.subheader("정밀조사표")
                col1, col2 = st.columns(2)
                with col1:
                    정밀_작업공정명 = st.text_input("작업공정명", key=f"정밀_작업공정명_{조사명}")
                with col2:
                    정밀_작업명 = st.text_input("작업명", key=f"정밀_작업명_{조사명}")
                
                # 사진 업로드 영역
                st.markdown("#### 사진")
                정밀_사진 = st.file_uploader(
                    "작업 사진 업로드",
                    type=['png', 'jpg', 'jpeg'],
                    accept_multiple_files=True,
                    key=f"정밀_사진_{조사명}"
                )
                if 정밀_사진:
                    cols = st.columns(3)
                    for photo_idx, photo in enumerate(정밀_사진):
                        with cols[photo_idx % 3]:
                            st.image(photo, caption=f"사진 {photo_idx+1}", use_column_width=True)
                
                st.markdown("---")
                
                # 작업별로 관련된 유해요인에 대한 원인분석
                st.markdown("#### ■ 작업별로 관련된 유해요인에 대한 원인분석")
                
                정밀_원인분석_data = []
                for i in range(7):
                    정밀_원인분석_data.append({
                        "작업분석 및 평가도구": "",
                        "분석결과": "",
                        "만점": ""
                    })
                
                정밀_원인분석_df = pd.DataFrame(정밀_원인분석_data)
                
                정밀_원인분석_config = {
                    "작업분석 및 평가도구": st.column_config.TextColumn("작업분석 및 평가도구", width=350),
                    "분석결과": st.column_config.TextColumn("분석결과", width=250),
                    "만점": st.column_config.TextColumn("만점", width=150)
                }
                
                정밀_원인분석_edited = st.data_editor(
                    정밀_원인분석_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config=정밀_원인분석_config,
                    num_rows="dynamic",
                    key=f"정밀_원인분석_{조사명}"
                )
                
                # 데이터 세션 상태에 저장
                st.session_state[f"정밀_원인분석_data_{조사명}"] = 정밀_원인분석_edited

# 6. 증상조사 분석 탭
with tabs[5]:
    st.title("근골격계 자기증상 분석")
    
    # 작업명 목록 가져오기
    작업명_목록 = get_작업명_목록()
    
    # 1. 기초현황
    st.subheader("1. 기초현황")
    
    # 작업명별 데이터 자동 생성
    기초현황_columns = ["작업명", "응답자(명)", "나이", "근속년수", "남자(명)", "여자(명)", "합계"]
    
    if 작업명_목록:
        # 작업명 목록을 기반으로 데이터 생성
        기초현황_data_rows = []
        for 작업명 in 작업명_목록:
            기초현황_data_rows.append([작업명, "", "평균(세)", "평균(년)", "", "", ""])
        기초현황_data = pd.DataFrame(기초현황_data_rows, columns=기초현황_columns)
    else:
        기초현황_data = pd.DataFrame(
            columns=기초현황_columns,
            data=[["", "", "평균(세)", "평균(년)", "", "", ""] for _ in range(3)]
        )
    
    기초현황_edited = st.data_editor(
        기초현황_data,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="기초현황_data"
    )
    
    # 2. 작업기간
    st.subheader("2. 작업기간")
    st.markdown("##### 현재 작업기간 / 이전 작업기간")
    
    작업기간_columns = ["작업명", "<1년", "<3년", "<5년", "≥5년", "무응답", "합계", "이전<1년", "이전<3년", "이전<5년", "이전≥5년", "이전무응답", "이전합계"]
    
    if 작업명_목록:
        # 작업명 목록을 기반으로 데이터 생성
        작업기간_data_rows = []
        for 작업명 in 작업명_목록:
            작업기간_data_rows.append([작업명] + [""] * 12)
        작업기간_data = pd.DataFrame(작업기간_data_rows, columns=작업기간_columns)
    else:
        작업기간_data = pd.DataFrame(
            columns=작업기간_columns,
            data=[[""] * 13 for _ in range(3)]
        )
    
    작업기간_edited = st.data_editor(
        작업기간_data,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="작업기간_data"
    )
    
    # 3. 육체적 부담정도
    st.subheader("3. 육체적 부담정도")
    육체적부담_columns = ["작업명", "전혀 힘들지 않음", "견딜만 함", "약간 힘듦", "힘듦", "매우 힘듦", "합계"]
    
    if 작업명_목록:
        # 작업명 목록을 기반으로 데이터 생성
        육체적부담_data_rows = []
        for 작업명 in 작업명_목록:
            육체적부담_data_rows.append([작업명] + [""] * 6)
        육체적부담_data = pd.DataFrame(육체적부담_data_rows, columns=육체적부담_columns)
    else:
        육체적부담_data = pd.DataFrame(
            columns=육체적부담_columns,
            data=[["", "", "", "", "", "", ""] for _ in range(3)]
        )
    
    육체적부담_edited = st.data_editor(
        육체적부담_data,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="육체적부담_data"
    )
    
    # 세션 상태에 저장
    st.session_state["기초현황_data_저장"] = 기초현황_edited
    st.session_state["작업기간_data_저장"] = 작업기간_edited
    st.session_state["육체적부담_data_저장"] = 육체적부담_edited
    
    # 4. 근골격계 통증 호소자 분포
    st.subheader("4. 근골격계 통증 호소자 분포")
    
    if 작업명_목록:
        # 컬럼 정의
        통증호소자_columns = ["작업명", "구분", "목", "어깨", "팔/팔꿈치", "손/손목/손가락", "허리", "다리/발", "전체"]
        
        # 데이터 생성
        통증호소자_data = []
        
        for 작업명 in 작업명_목록:
            # 각 작업명에 대해 정상, 관리대상자, 통증호소자 3개 행 추가
            통증호소자_data.append([작업명, "정상", "", "", "", "", "", "", ""])
            통증호소자_data.append(["", "관리대상자", "", "", "", "", "", "", ""])
            통증호소자_data.append(["", "통증호소자", "", "", "", "", "", "", ""])
        
        통증호소자_df = pd.DataFrame(통증호소자_data, columns=통증호소자_columns)
        
        # 컬럼 설정
        column_config = {
            "작업명": st.column_config.TextColumn("작업명", disabled=True, width=150),
            "구분": st.column_config.TextColumn("구분", disabled=True, width=100),
            "목": st.column_config.TextColumn("목", width=80),
            "어깨": st.column_config.TextColumn("어깨", width=80),
            "팔/팔꿈치": st.column_config.TextColumn("팔/팔꿈치", width=100),
            "손/손목/손가락": st.column_config.TextColumn("손/손목/손가락", width=120),
            "허리": st.column_config.TextColumn("허리", width=80),
            "다리/발": st.column_config.TextColumn("다리/발", width=80),
            "전체": st.column_config.TextColumn("전체", width=80)
        }
        
        통증호소자_edited = st.data_editor(
            통증호소자_df,
            hide_index=True,
            use_container_width=True,
            column_config=column_config,
            key="통증호소자_data_editor",
            disabled=["작업명", "구분"]
        )
        
        # 세션 상태에 저장
        st.session_state["통증호소자_data_저장"] = 통증호소자_edited
    else:
        st.info("체크리스트에 작업명을 입력하면 자동으로 표가 생성됩니다.")
        
        # 빈 데이터프레임 표시
        통증호소자_columns = ["작업명", "구분", "목", "어깨", "팔/팔꿈치", "손/손목/손가락", "허리", "다리/발", "전체"]
        빈_df = pd.DataFrame(columns=통증호소자_columns)
        st.dataframe(빈_df, use_container_width=True)

# 7. 작업환경개선계획서 탭
with tabs[6]:
    st.title("작업환경개선계획서")
    
    # 컬럼 정의
    개선계획_columns = [
        "공정명",
        "작업명",
        "단위작업명",
        "문제점(유해요인의 원인)",
        "근로자의견",
        "개선방안",
        "추진일정",
        "개선비용",
        "개선우선순위"
    ]
    
    # 세션 상태에 개선계획 데이터가 없으면 초기화
    if "개선계획_data_저장" not in st.session_state or st.session_state["개선계획_data_저장"].empty:
        # 체크리스트 데이터 기반으로 초기 데이터 생성
        if not st.session_state["checklist_df"].empty:
            개선계획_data_rows = []
            for _, row in st.session_state["checklist_df"].iterrows():
                if row["작업명"] and row["단위작업명"]:
                    # 부담작업이 있는 경우만 추가
                    부담작업_있음 = False
                    for i in range(1, 12):
                        if row[f"{i}호"] in ["O(해당)", "△(잠재위험)"]:
                            부담작업_있음 = True
                            break
                    
                    if 부담작업_있음:
                        개선계획_data_rows.append([
                            row["작업명"],  # 공정명에도 작업명 사용
                            row["작업명"],
                            row["단위작업명"],
                            "",  # 문제점
                            "",  # 근로자의견
                            "",  # 개선방안
                            "",  # 추진일정
                            "",  # 개선비용
                            ""   # 개선우선순위
                        ])
            
            # 데이터가 있으면 사용, 없으면 빈 행 5개
            if 개선계획_data_rows:
                개선계획_data = pd.DataFrame(개선계획_data_rows, columns=개선계획_columns)
            else:
                개선계획_data = pd.DataFrame(
                    columns=개선계획_columns,
                    data=[["", "", "", "", "", "", "", "", ""] for _ in range(5)]
                )
        else:
            # 초기 데이터 (빈 행 5개)
            개선계획_data = pd.DataFrame(
                columns=개선계획_columns,
                data=[["", "", "", "", "", "", "", "", ""] for _ in range(5)]
            )
        
        st.session_state["개선계획_data_저장"] = 개선계획_data
    
    # 컬럼 설정
    개선계획_config = {
        "공정명": st.column_config.TextColumn("공정명", width=100),
        "작업명": st.column_config.TextColumn("작업명", width=100),
        "단위작업명": st.column_config.TextColumn("단위작업명", width=120),
        "문제점(유해요인의 원인)": st.column_config.TextColumn("문제점(유해요인의 원인)", width=200),
        "근로자의견": st.column_config.TextColumn("근로자의견", width=150),
        "개선방안": st.column_config.TextColumn("개선방안", width=200),
        "추진일정": st.column_config.TextColumn("추진일정", width=100),
        "개선비용": st.column_config.TextColumn("개선비용", width=100),
        "개선우선순위": st.column_config.TextColumn("개선우선순위", width=120)
    }
    
    # 데이터 편집기
    개선계획_edited = st.data_editor(
        st.session_state["개선계획_data_저장"],
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        column_config=개선계획_config,
        key="개선계획_data"
    )
    
    # 세션 상태에 저장
    st.session_state["개선계획_data_저장"] = 개선계획_edited
    
    # 도움말
    with st.expander("ℹ️ 작성 도움말"):
        st.markdown("""
        - **공정명**: 해당 작업이 속한 공정명
        - **작업명**: 개선이 필요한 작업명
        - **단위작업명**: 구체적인 단위작업명
        - **문제점**: 유해요인의 구체적인 원인
        - **근로자의견**: 현장 근로자의 개선 의견
        - **개선방안**: 구체적인 개선 방법
        - **추진일정**: 개선 예정 시기
        - **개선비용**: 예상 소요 비용
        - **개선우선순위**: 종합점수/중점수/중상호소여부를 고려한 우선순위
        """)
    
    # 행 추가/삭제 버튼
    col1, col2, col3 = st.columns([8, 1, 1])
    with col2:
        if st.button("➕ 행 추가", key="개선계획_행추가", use_container_width=True):
            new_row = pd.DataFrame([["", "", "", "", "", "", "", "", ""]], columns=개선계획_columns)
            st.session_state["개선계획_data_저장"] = pd.concat([st.session_state["개선계획_data_저장"], new_row], ignore_index=True)
            st.rerun()
    with col3:
        if st.button("➖ 마지막 행 삭제", key="개선계획_행삭제", use_container_width=True):
            if len(st.session_state["개선계획_data_저장"]) > 0:
                st.session_state["개선계획_data_저장"] = st.session_state["개선계획_data_저장"].iloc[:-1]
                st.rerun()
    
    # 전체 보고서 다운로드
    st.markdown("---")
    st.subheader("📥 전체 보고서 다운로드")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # 엑셀 다운로드 버튼
        if st.button("📊 엑셀 파일로 다운로드", use_container_width=True):
            try:
                output = BytesIO()
                
                # 작업명 목록 다시 가져오기
                작업명_목록_다운로드 = get_작업명_목록()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # 사업장 개요 정보
                    overview_data = {
                        "항목": ["사업장명", "소재지", "업종", "예비조사일", "본조사일", "수행기관", "성명"],
                        "내용": [
                            st.session_state.get("사업장명", ""),
                            st.session_state.get("소재지", ""),
                            st.session_state.get("업종", ""),
                            str(st.session_state.get("예비조사", "")),
                            str(st.session_state.get("본조사", "")),
                            st.session_state.get("수행기관", ""),
                            st.session_state.get("성명", "")
                        ]
                    }
                    overview_df = pd.DataFrame(overview_data)
                    overview_df.to_excel(writer, sheet_name='사업장개요', index=False)
                    
                    # 체크리스트
                    if "checklist_df" in st.session_state and not st.session_state["checklist_df"].empty:
                        st.session_state["checklist_df"].to_excel(writer, sheet_name='체크리스트', index=False)
                    
                    # 유해요인조사표 데이터 저장 (작업명별로)
                    for 작업명 in 작업명_목록_다운로드:
                        조사표_data = []
                        
                        # 조사개요
                        조사표_data.append(["조사개요"])
                        조사표_data.append(["조사일시", st.session_state.get(f"조사일시_{작업명}", "")])
                        조사표_data.append(["부서명", st.session_state.get(f"부서명_{작업명}", "")])
                        조사표_data.append(["조사자", st.session_state.get(f"조사자_{작업명}", "")])
                        조사표_data.append(["작업공정명", st.session_state.get(f"작업공정명_{작업명}", "")])
                        조사표_data.append(["작업명", st.session_state.get(f"작업명_{작업명}", "")])
                        조사표_data.append([])  # 빈 행
                        
                        # 작업장 상황조사
                        조사표_data.append(["작업장 상황조사"])
                        조사표_data.append(["항목", "상태", "세부사항"])
                        
                        for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
                            상태 = st.session_state.get(f"{항목}_상태_{작업명}", "변화없음")
                            세부사항 = ""
                            if 상태 == "감소":
                                세부사항 = st.session_state.get(f"{항목}_감소_시작_{작업명}", "")
                            elif 상태 == "증가":
                                세부사항 = st.session_state.get(f"{항목}_증가_시작_{작업명}", "")
                            elif 상태 == "기타":
                                세부사항 = st.session_state.get(f"{항목}_기타_내용_{작업명}", "")
                            
                            조사표_data.append([항목, 상태, 세부사항])
                        
                        if 조사표_data:
                            조사표_df = pd.DataFrame(조사표_data)
                            sheet_name = f'유해요인_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                            조사표_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # 각 작업별 데이터 저장
                    for 작업명 in 작업명_목록_다운로드:
                        # 작업조건조사 데이터 저장
                        data_key = f"작업조건_data_{작업명}"
                        if data_key in st.session_state:
                            작업_df = st.session_state[data_key]
                            if isinstance(작업_df, pd.DataFrame) and not 작업_df.empty:
                                export_df = 작업_df.copy()
                                
                                # 총점 계산
                                for idx in range(len(export_df)):
                                    export_df.at[idx, "총점"] = calculate_total_score(export_df.iloc[idx])
                                
                                # 시트 이름 정리 (특수문자 제거)
                                sheet_name = f'작업조건_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                                export_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 3단계 유해요인평가 데이터 저장
                        평가_작업명 = st.session_state.get(f"3단계_작업명_{작업명}", 작업명)
                        평가_근로자수 = st.session_state.get(f"3단계_근로자수_{작업명}", "")
                        
                        평가_data = {
                            "작업명": [평가_작업명],
                            "근로자수": [평가_근로자수]
                        }
                        
                        # 사진 설명 추가
                        사진개수 = st.session_state.get(f"사진개수_{작업명}", 3)
                        for i in range(사진개수):
                            설명 = st.session_state.get(f"사진_{i+1}_설명_{작업명}", "")
                            평가_data[f"사진{i+1}_설명"] = [설명]
                        
                        if 평가_작업명 or 평가_근로자수:
                            평가_df = pd.DataFrame(평가_data)
                            sheet_name = f'유해요인평가_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                            평가_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 원인분석 데이터 저장 (개선된 버전)
                        원인분석_key = f"원인분석_항목_{작업명}"
                        if 원인분석_key in st.session_state:
                            원인분석_data = []
                            for item in st.session_state[원인분석_key]:
                                if item.get("단위작업명") or item.get("유형"):
                                    원인분석_data.append(item)
                            
                            if 원인분석_data:
                                원인분석_df = pd.DataFrame(원인분석_data)
                                sheet_name = f'원인분석_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                                원인분석_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 정밀조사 데이터 저장 (조사명별로)
                    if "정밀조사_목록" in st.session_state and st.session_state["정밀조사_목록"]:
                        for 조사명 in st.session_state["정밀조사_목록"]:
                            정밀_data_rows = []
                            
                            # 기본 정보
                            정밀_data_rows.append(["작업공정명", st.session_state.get(f"정밀_작업공정명_{조사명}", "")])
                            정밀_data_rows.append(["작업명", st.session_state.get(f"정밀_작업명_{조사명}", "")])
                            정밀_data_rows.append([])  # 빈 행
                            정밀_data_rows.append(["작업별로 관련된 유해요인에 대한 원인분석"])
                            정밀_data_rows.append(["작업분석 및 평가도구", "분석결과", "만점"])
                            
                            # 원인분석 데이터
                            원인분석_key = f"정밀_원인분석_data_{조사명}"
                            if 원인분석_key in st.session_state:
                                원인분석_df = st.session_state[원인분석_key]
                                for _, row in 원인분석_df.iterrows():
                                    if row.get("작업분석 및 평가도구", "") or row.get("분석결과", "") or row.get("만점", ""):
                                        정밀_data_rows.append([
                                            row.get("작업분석 및 평가도구", ""),
                                            row.get("분석결과", ""),
                                            row.get("만점", "")
                                        ])
                            
                            if len(정밀_data_rows) > 5:  # 헤더 이후에 데이터가 있는 경우만
                                정밀_sheet_df = pd.DataFrame(정밀_data_rows)
                                sheet_name = 조사명.replace('/', '_').replace('\\', '_')[:31]
                                정밀_sheet_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # 증상조사 분석 데이터 저장
                    if "기초현황_data_저장" in st.session_state:
                        기초현황_df = st.session_state["기초현황_data_저장"]
                        if not 기초현황_df.empty:
                            기초현황_df.to_excel(writer, sheet_name="증상조사_기초현황", index=False)

                    if "작업기간_data_저장" in st.session_state:
                        작업기간_df = st.session_state["작업기간_data_저장"]
                        if not 작업기간_df.empty:
                            작업기간_df.to_excel(writer, sheet_name="증상조사_작업기간", index=False)

                    if "육체적부담_data_저장" in st.session_state:
                        육체적부담_df = st.session_state["육체적부담_data_저장"]
                        if not 육체적부담_df.empty:
                            육체적부담_df.to_excel(writer, sheet_name="증상조사_육체적부담", index=False)

                    if "통증호소자_data_저장" in st.session_state:
                        통증호소자_df = st.session_state["통증호소자_data_저장"]
                        if isinstance(통증호소자_df, pd.DataFrame) and not 통증호소자_df.empty:
                            통증호소자_df.to_excel(writer, sheet_name="증상조사_통증호소자", index=False)
                    
                    # 작업환경개선계획서 데이터 저장
                    if "개선계획_data_저장" in st.session_state:
                        개선계획_df = st.session_state["개선계획_data_저장"]
                        if not 개선계획_df.empty:
                            # 빈 행 제거 (모든 컬럼이 빈 행 제외)
                            개선계획_df_clean = 개선계획_df[개선계획_df.astype(str).ne('').any(axis=1)]
                            if not 개선계획_df_clean.empty:
                                개선계획_df_clean.to_excel(writer, sheet_name="작업환경개선계획서", index=False)
                    
                output.seek(0)
                st.download_button(
                    label="📥 엑셀 다운로드",
                    data=output,
                    file_name=f"근골격계_유해요인조사_{st.session_state.get('workplace', '')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"엑셀 파일 생성 중 오류가 발생했습니다: {str(e)}")
                st.info("데이터를 입력한 후 다시 시도해주세요.")
    
    with col2:
        # PDF 보고서 생성 버튼
        if PDF_AVAILABLE:
            if st.button("📄 PDF 보고서 생성", use_container_width=True):
                try:
                    # 한글 폰트 설정 - 나눔고딕 우선
                    font_paths = [
                        "C:/Windows/Fonts/NanumGothic.ttf",
                        "C:/Windows/Fonts/NanumBarunGothic.ttf",
                        "C:/Windows/Fonts/malgun.ttf",
                        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",  # Linux
                        "/System/Library/Fonts/Supplemental/NanumGothic.ttf"  # Mac
                    ]
                    
                    font_registered = False
                    for font_path in font_paths:
                        if os.path.exists(font_path):
                            if "NanumGothic" in font_path:
                                pdfmetrics.registerFont(TTFont('NanumGothic', font_path))
                                font_name = 'NanumGothic'
                            elif "NanumBarunGothic" in font_path:
                                pdfmetrics.registerFont(TTFont('NanumBarunGothic', font_path))
                                font_name = 'NanumBarunGothic'
                            else:
                                pdfmetrics.registerFont(TTFont('Malgun', font_path))
                                font_name = 'Malgun'
                            font_registered = True
                            break
                    
                    if not font_registered:
                        font_name = 'Helvetica'
                    
                    # PDF 생성
                    pdf_buffer = BytesIO()
                    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
                    story = []
                    
                    # 스타일 설정 - 글꼴 크기 증가
                    styles = getSampleStyleSheet()
                    title_style = ParagraphStyle(
                        'CustomTitle',
                        parent=styles['Heading1'],
                        fontSize=28,  # 24에서 28로 증가
                        textColor=colors.HexColor('#1f4788'),
                        alignment=TA_CENTER,
                        fontName=font_name,
                        spaceAfter=30
                    )
                    
                    heading_style = ParagraphStyle(
                        'CustomHeading',
                        parent=styles['Heading2'],
                        fontSize=18,  # 16에서 18로 증가
                        textColor=colors.HexColor('#2e5090'),
                        fontName=font_name,
                        spaceAfter=12
                    )
                    
                    subheading_style = ParagraphStyle(
                        'CustomSubHeading',
                        parent=styles['Heading3'],
                        fontSize=14,  # 새로 추가
                        textColor=colors.HexColor('#3a5fa0'),
                        fontName=font_name,
                        spaceAfter=10
                    )
                    
                    normal_style = ParagraphStyle(
                        'CustomNormal',
                        parent=styles['Normal'],
                        fontSize=12,  # 10에서 12로 증가
                        fontName=font_name,
                        leading=14
                    )
                    
                    # 제목 페이지
                    story.append(Spacer(1, 1.5*inch))
                    story.append(Paragraph("근골격계 유해요인조사 보고서", title_style))
                    story.append(Spacer(1, 0.5*inch))
                    
                    # 사업장 정보
                    if st.session_state.get("사업장명"):
                        사업장정보 = f"""
                        <para align="center" fontSize="14">
                        <b>사업장명:</b> {st.session_state.get("사업장명", "")}<br/>
                        <b>작업현장:</b> {st.session_state.get("workplace", "")}<br/>
                        <b>조사일:</b> {datetime.now().strftime('%Y년 %m월 %d일')}
                        </para>
                        """
                        story.append(Paragraph(사업장정보, normal_style))
                    
                    story.append(PageBreak())
                    
                    # 1. 사업장 개요
                    story.append(Paragraph("1. 사업장 개요", heading_style))
                    
                    사업장_data = [
                        ["항목", "내용"],
                        ["사업장명", st.session_state.get("사업장명", "")],
                        ["소재지", st.session_state.get("소재지", "")],
                        ["업종", st.session_state.get("업종", "")],
                        ["예비조사일", str(st.session_state.get("예비조사", ""))],
                        ["본조사일", str(st.session_state.get("본조사", ""))],
                        ["수행기관", st.session_state.get("수행기관", "")],
                        ["담당자", st.session_state.get("성명", "")]
                    ]
                    
                    t = Table(사업장_data, colWidths=[2*inch, 4*inch])
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, -1), font_name),
                        ('FONTSIZE', (0, 0), (-1, -1), 12),  # 10에서 12로 증가
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
                        ('BACKGROUND', (0, 1), (0, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 0.5*inch))
                    
                    # PDF 생성 (나머지 부분 생략 - 기존 코드와 동일)
                    doc.build(story)
                    pdf_buffer.seek(0)
                    
                    # 다운로드 버튼
                    st.download_button(
                        label="📥 PDF 다운로드",
                        data=pdf_buffer,
                        file_name=f"근골격계유해요인조사보고서_{st.session_state.get('workplace', '')}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf"
                    )
                    
                    st.success("PDF 보고서가 생성되었습니다!")
                    
                except Exception as e:
                    error_message = "PDF 생성 중 오류가 발생했습니다: " + str(e)
                    st.error(error_message)
                    st.info("reportlab 라이브러리를 설치해주세요: pip install reportlab")
        else:
            st.info("PDF 생성 기능을 사용하려면 reportlab 라이브러리를 설치하세요: pip install reportlab")
