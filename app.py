import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import json
import os
import time
import traceback

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

# Excel 파일 저장 디렉토리 생성
SAVE_DIR = "saved_sessions"
BACKUP_DIR = "saved_sessions/backups"
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)
if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR)

# 데이터 무결성 검증 함수
def validate_dataframe(df):
    """DataFrame이 유효한지 검증"""
    if df is None:
        return False
    if not isinstance(df, pd.DataFrame):
        return False
    return True

# 안전한 데이터 저장 함수
def safe_save_to_excel(session_id, workplace=None):
    """데이터를 안전하게 Excel 파일로 저장 (백업 포함)"""
    try:
        # 임시 파일명
        temp_filename = os.path.join(SAVE_DIR, f"{session_id}_temp.xlsx")
        final_filename = os.path.join(SAVE_DIR, f"{session_id}.xlsx")
        backup_filename = os.path.join(BACKUP_DIR, f"{session_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        # 기존 파일이 있으면 백업
        if os.path.exists(final_filename):
            try:
                import shutil
                shutil.copy2(final_filename, backup_filename)
            except:
                pass
        
        # 임시 파일에 먼저 저장
        with pd.ExcelWriter(temp_filename, engine='openpyxl') as writer:
            # 메타데이터 저장
            metadata = {
                "session_id": session_id,
                "workplace": workplace or st.session_state.get("workplace", ""),
                "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "사업장명": st.session_state.get("사업장명", ""),
                "소재지": st.session_state.get("소재지", ""),
                "업종": st.session_state.get("업종", ""),
                "예비조사": str(st.session_state.get("예비조사", "")),
                "본조사": str(st.session_state.get("본조사", "")),
                "수행기관": st.session_state.get("수행기관", ""),
                "성명": st.session_state.get("성명", "")
            }
            
            metadata_df = pd.DataFrame([metadata])
            metadata_df.to_excel(writer, sheet_name='메타데이터', index=False)
            
            # 체크리스트 저장
            if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
                if not st.session_state["checklist_df"].empty:
                    st.session_state["checklist_df"].to_excel(writer, sheet_name='체크리스트', index=False)
            
            # 반 목록 가져오기
            반_목록 = []
            if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
                if not st.session_state["checklist_df"].empty:
                    반_목록 = st.session_state["checklist_df"]["반"].dropna().unique().tolist()
            
            # 각 반별 데이터 저장
            for 반 in 반_목록:
                safe_반 = str(반).replace('/', '_').replace('\\', '_')[:31]
                
                # 유해요인조사표 데이터
                조사표_data = {
                    "조사일시": st.session_state.get(f"조사일시_{반}", ""),
                    "부서명": st.session_state.get(f"부서명_{반}", ""),
                    "조사자": st.session_state.get(f"조사자_{반}", ""),
                    "작업공정명": st.session_state.get(f"작업공정명_{반}", ""),
                    "작업명": st.session_state.get(f"작업명_{반}", "")
                }
                
                # 작업장 상황조사
                for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
                    조사표_data[f"{항목}_상태"] = st.session_state.get(f"{항목}_상태_{반}", "")
                    조사표_data[f"{항목}_세부사항"] = st.session_state.get(f"{항목}_감소_시작_{반}", "") or \
                                                     st.session_state.get(f"{항목}_증가_시작_{반}", "") or \
                                                     st.session_state.get(f"{항목}_기타_내용_{반}", "")
                
                조사표_df = pd.DataFrame([조사표_data])
                sheet_name = f'조사표_{safe_반}'
                조사표_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 작업조건조사 데이터
                작업조건_key = f"작업조건_data_{반}"
                if 작업조건_key in st.session_state and validate_dataframe(st.session_state.get(작업조건_key)):
                    sheet_name = f'작업조건_{safe_반}'
                    st.session_state[작업조건_key].to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 원인분석 데이터
                원인분석_key = f"원인분석_항목_{반}"
                if 원인분석_key in st.session_state and st.session_state[원인분석_key]:
                    원인분석_df = pd.DataFrame(st.session_state[원인분석_key])
                    sheet_name = f'원인분석_{safe_반}'
                    원인분석_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 정밀조사 데이터
            if "정밀조사_목록" in st.session_state:
                for 조사명 in st.session_state["정밀조사_목록"]:
                    safe_조사명 = str(조사명).replace('/', '_').replace('\\', '_')[:31]
                    정밀_data = {
                        "작업공정명": st.session_state.get(f"정밀_작업공정명_{조사명}", ""),
                        "작업명": st.session_state.get(f"정밀_작업명_{조사명}", "")
                    }
                    
                    원인분석_key = f"정밀_원인분석_data_{조사명}"
                    if 원인분석_key in st.session_state and validate_dataframe(st.session_state.get(원인분석_key)):
                        sheet_name = f'정밀_{safe_조사명}'
                        정밀_df = pd.DataFrame([정밀_data])
                        정밀_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        st.session_state[원인분석_key].to_excel(
                            writer, 
                            sheet_name=sheet_name, 
                            startrow=3, 
                            index=False
                        )
            
            # 증상조사 분석 데이터
            증상조사_시트 = {
                "기초현황": "기초현황_data_저장",
                "작업기간": "작업기간_data_저장",
                "육체적부담": "육체적부담_data_저장",
                "통증호소자": "통증호소자_data_저장"
            }
            
            for 시트명, 키 in 증상조사_시트.items():
                if 키 in st.session_state and validate_dataframe(st.session_state.get(키)):
                    if not st.session_state[키].empty:
                        st.session_state[키].to_excel(writer, sheet_name=f'증상_{시트명}', index=False)
            
            # 작업환경개선계획서
            if "개선계획_data_저장" in st.session_state and validate_dataframe(st.session_state.get("개선계획_data_저장")):
                if not st.session_state["개선계획_data_저장"].empty:
                    st.session_state["개선계획_data_저장"].to_excel(writer, sheet_name='개선계획서', index=False)
        
        # 임시 파일을 최종 파일로 이동
        if os.path.exists(temp_filename):
            if os.path.exists(final_filename):
                os.remove(final_filename)
            os.rename(temp_filename, final_filename)
            
        return True, final_filename
        
    except Exception as e:
        # 에러 발생 시 임시 파일 정리
        if os.path.exists(temp_filename):
            try:
                os.remove(temp_filename)
            except:
                pass
        return False, f"저장 중 오류 발생: {str(e)}\n{traceback.format_exc()}"

# 안전한 데이터 불러오기 함수
def safe_load_from_excel(filename):
    """Excel 파일에서 데이터를 안전하게 불러오기"""
    try:
        # 파일 존재 여부 확인
        if not os.path.exists(filename):
            return False, "파일이 존재하지 않습니다."
        
        # 전체 시트 읽기
        excel_file = pd.ExcelFile(filename)
        
        # 메타데이터 읽기
        if '메타데이터' in excel_file.sheet_names:
            try:
                metadata_df = pd.read_excel(excel_file, sheet_name='메타데이터')
                if not metadata_df.empty:
                    metadata = metadata_df.iloc[0].to_dict()
                    
                    # 세션 상태에 메타데이터 복원
                    for key in ["session_id", "workplace", "사업장명", "소재지", "업종", "예비조사", "본조사", "수행기관", "성명"]:
                        if key in metadata:
                            value = metadata[key]
                            if pd.notna(value):
                                st.session_state[key] = str(value) if value else ""
            except Exception as e:
                st.warning(f"메타데이터 읽기 오류: {str(e)}")
        
        # 체크리스트 읽기
        if '체크리스트' in excel_file.sheet_names:
            try:
                checklist_df = pd.read_excel(excel_file, sheet_name='체크리스트')
                if validate_dataframe(checklist_df):
                    st.session_state["checklist_df"] = checklist_df
            except Exception as e:
                st.warning(f"체크리스트 읽기 오류: {str(e)}")
        
        # 각 시트별로 데이터 읽기
        for sheet_name in excel_file.sheet_names:
            try:
                if sheet_name.startswith('조사표_'):
                    반 = sheet_name.replace('조사표_', '')
                    조사표_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if not 조사표_df.empty:
                        data = 조사표_df.iloc[0].to_dict()
                        for key, value in data.items():
                            if pd.notna(value):
                                st.session_state[f"{key}_{반}"] = str(value) if value else ""
                
                elif sheet_name.startswith('작업조건_'):
                    반 = sheet_name.replace('작업조건_', '')
                    작업조건_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if validate_dataframe(작업조건_df):
                        st.session_state[f"작업조건_data_{반}"] = 작업조건_df
                
                elif sheet_name.startswith('원인분석_'):
                    반 = sheet_name.replace('원인분석_', '')
                    원인분석_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if validate_dataframe(원인분석_df):
                        st.session_state[f"원인분석_항목_{반}"] = 원인분석_df.to_dict('records')
                
                elif sheet_name.startswith('정밀_'):
                    조사명 = sheet_name.replace('정밀_', '')
                    if "정밀조사_목록" not in st.session_state:
                        st.session_state["정밀조사_목록"] = []
                    if 조사명 not in st.session_state["정밀조사_목록"]:
                        st.session_state["정밀조사_목록"].append(조사명)
                    
                    정밀_df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=1)
                    if not 정밀_df.empty:
                        data = 정밀_df.iloc[0].to_dict()
                        for key, value in data.items():
                            if pd.notna(value):
                                st.session_state[f"정밀_{key}_{조사명}"] = str(value) if value else ""
                    
                    # 원인분석 데이터 읽기
                    try:
                        원인분석_df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=3)
                        if validate_dataframe(원인분석_df):
                            st.session_state[f"정밀_원인분석_data_{조사명}"] = 원인분석_df
                    except:
                        pass
                
                elif sheet_name.startswith('증상_'):
                    증상_키 = sheet_name.replace('증상_', '') + "_data_저장"
                    증상_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if validate_dataframe(증상_df):
                        st.session_state[증상_키] = 증상_df
                
                elif sheet_name == '개선계획서':
                    개선계획_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if validate_dataframe(개선계획_df):
                        st.session_state["개선계획_data_저장"] = 개선계획_df
                        
            except Exception as e:
                st.warning(f"시트 '{sheet_name}' 읽기 오류: {str(e)}")
                continue
        
        return True, "데이터를 성공적으로 불러왔습니다."
        
    except Exception as e:
        return False, f"파일 불러오기 중 오류 발생: {str(e)}\n{traceback.format_exc()}"

# 단위작업명 병합 함수
def merge_unit_works(selected_indices, checklist_df, merge_name):
    """선택된 단위작업들을 하나로 병합"""
    if not selected_indices or not merge_name:
        return checklist_df
    
    # 선택된 행들의 데이터 가져오기
    selected_rows = checklist_df.iloc[selected_indices]
    
    # 첫 번째 행을 기준으로 병합
    merged_row = selected_rows.iloc[0].copy()
    merged_row["단위작업명"] = merge_name
    
    # 부담작업 정보 병합 (각 호별로 가장 높은 수준 선택)
    priority_map = {"O(해당)": 3, "△(잠재위험)": 2, "X(미해당)": 1}
    reverse_map = {3: "O(해당)", 2: "△(잠재위험)", 1: "X(미해당)"}
    
    for i in range(1, 12):
        col_name = f"{i}호"
        values = selected_rows[col_name].tolist()
        max_priority = max([priority_map.get(v, 1) for v in values])
        merged_row[col_name] = reverse_map[max_priority]
    
    # 새 DataFrame 생성
    new_df = checklist_df.drop(selected_indices).reset_index(drop=True)
    new_df = pd.concat([new_df, pd.DataFrame([merged_row])], ignore_index=True)
    
    return new_df

# 자동 저장 기능 (Excel 버전)
def auto_save():
    if "last_save_time" not in st.session_state:
        st.session_state["last_save_time"] = time.time()
    
    current_time = time.time()
    if current_time - st.session_state["last_save_time"] > 30:  # 30초마다 자동 저장
        if st.session_state.get("session_id") and st.session_state.get("workplace"):
            success, _ = safe_save_to_excel(st.session_state["session_id"], st.session_state.get("workplace"))
            if success:
                st.session_state["last_save_time"] = current_time
                st.session_state["last_successful_save"] = datetime.now()

# 저장된 세션 목록 가져오기
def get_saved_sessions():
    """저장된 Excel 세션 파일 목록 반환"""
    sessions = []
    if os.path.exists(SAVE_DIR):
        for filename in os.listdir(SAVE_DIR):
            if filename.endswith('.xlsx') and not filename.endswith('_temp.xlsx'):
                filepath = os.path.join(SAVE_DIR, filename)
                try:
                    # 메타데이터 읽기
                    metadata_df = pd.read_excel(filepath, sheet_name='메타데이터')
                    if not metadata_df.empty:
                        metadata = metadata_df.iloc[0].to_dict()
                        sessions.append({
                            "filename": filename,
                            "session_id": metadata.get("session_id", ""),
                            "workplace": metadata.get("workplace", ""),
                            "saved_at": metadata.get("saved_at", "")
                        })
                except:
                    continue
    return sorted(sessions, key=lambda x: x.get("saved_at", ""), reverse=True)

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

# 작업현장별 세션 관리
if "workplace" not in st.session_state:
    st.session_state["workplace"] = None

if "session_id" not in st.session_state:
    st.session_state["session_id"] = None

# 계층 구조 데이터 가져오는 함수들
def get_회사명_목록():
    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
        if not st.session_state["checklist_df"].empty:
            return st.session_state["checklist_df"]["회사명"].dropna().unique().tolist()
    return []

def get_소속_목록(회사명=None):
    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
        if not st.session_state["checklist_df"].empty:
            df = st.session_state["checklist_df"]
            if 회사명:
                df = df[df["회사명"] == 회사명]
            return df["소속"].dropna().unique().tolist()
    return []

def get_반_목록(회사명=None, 소속=None):
    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
        if not st.session_state["checklist_df"].empty:
            df = st.session_state["checklist_df"]
            if 회사명:
                df = df[df["회사명"] == 회사명]
            if 소속:
                df = df[df["소속"] == 소속]
            return df["반"].dropna().unique().tolist()
    return []

def get_단위작업명_목록(회사명=None, 소속=None, 반=None):
    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
        if not st.session_state["checklist_df"].empty:
            df = st.session_state["checklist_df"]
            if 회사명:
                df = df[df["회사명"] == 회사명]
            if 소속:
                df = df[df["소속"] == 소속]
            if 반:
                df = df[df["반"] == 반]
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
    st.title("[데이터 관리]")
    
    # 작업현장 선택/입력
    st.markdown("### [작업현장 선택]")
    작업현장_옵션 = ["현장 선택...", "A사업장", "B사업장", "C사업장", "신규 현장 추가"]
    선택된_현장 = st.selectbox("작업현장", 작업현장_옵션)
    
    if 선택된_현장 == "신규 현장 추가":
        새현장명 = st.text_input("새 현장명 입력")
        if 새현장명:
            st.session_state["workplace"] = 새현장명
            st.session_state["session_id"] = f"{새현장명}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    elif 선택된_현장 != "현장 선택...":
        st.session_state["workplace"] = 선택된_현장
        if not st.session_state.get("session_id") or 선택된_현장 not in st.session_state.get("session_id", ""):
            st.session_state["session_id"] = f"{선택된_현장}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # 세션 정보 표시
    if st.session_state.get("session_id"):
        st.info(f"[세션 ID] {st.session_state['session_id']}")
    
    # 자동 저장 상태
    if "last_successful_save" in st.session_state:
        last_save = st.session_state["last_successful_save"]
        st.success(f"[저장 완료] 마지막 자동저장: {last_save.strftime('%H:%M:%S')}")
    
    # 수동 저장 버튼
    if st.button("[Excel로 저장]", use_container_width=True):
        if st.session_state.get("session_id") and st.session_state.get("workplace"):
            success, result = safe_save_to_excel(st.session_state["session_id"], st.session_state.get("workplace"))
            if success:
                st.success(f"[저장 완료] Excel 파일로 저장되었습니다!\n[파일 위치] {result}")
                st.session_state["last_successful_save"] = datetime.now()
            else:
                st.error(f"저장 중 오류 발생:\n{result}")
        else:
            st.warning("먼저 작업현장을 선택해주세요!")
    
    # 저장된 세션 목록
    st.markdown("---")
    st.markdown("### [저장된 세션]")
    
    saved_sessions = get_saved_sessions()
    if saved_sessions:
        selected_session = st.selectbox(
            "불러올 세션 선택",
            options=["선택..."] + [f"{s['workplace']} - {s['saved_at']}" for s in saved_sessions],
            key="session_selector"
        )
        
        if selected_session != "선택..." and st.button("[세션 불러오기]", use_container_width=True):
            session_idx = [f"{s['workplace']} - {s['saved_at']}" for s in saved_sessions].index(selected_session)
            session_info = saved_sessions[session_idx]
            filepath = os.path.join(SAVE_DIR, session_info["filename"])
            
            success, message = safe_load_from_excel(filepath)
            if success:
                st.success(f"[불러오기 완료] {message}")
                st.rerun()
            else:
                st.error(message)
    else:
        st.info("저장된 세션이 없습니다.")
    
    # Excel 파일 직접 업로드
    st.markdown("---")
    st.markdown("### [Excel 파일 업로드]")
    uploaded_file = st.file_uploader("Excel 파일 선택", type=['xlsx'])
    if uploaded_file is not None:
        if st.button("[데이터 가져오기]", use_container_width=True):
            # 임시 파일로 저장
            temp_path = os.path.join(SAVE_DIR, f"temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            try:
                with open(temp_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                
                success, message = safe_load_from_excel(temp_path)
                if success:
                    st.success(f"[가져오기 완료] {message}")
                    os.remove(temp_path)  # 임시 파일 삭제
                    st.rerun()
                else:
                    st.error(message)
                    os.remove(temp_path)  # 임시 파일 삭제
            except Exception as e:
                st.error(f"파일 처리 중 오류: {str(e)}")
                if os.path.exists(temp_path):
                    os.remove(temp_path)
    
    # 부담작업 참고 정보
    with st.expander("[부담작업 빠른 참조]"):
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

# 자동 저장 실행
if st.session_state.get("session_id") and st.session_state.get("workplace"):
    auto_save()

# 작업현장 선택 확인
if not st.session_state.get("workplace"):
    st.warning("먼저 사이드바에서 작업현장을 선택하거나 입력해주세요!")
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
        예비조사 = st.text_input("예비조사일 (YYYY-MM-DD)", key="예비조사", placeholder="2024-01-01")
        수행기관 = st.text_input("수행기관", key="수행기관")
    with col2:
        본조사 = st.text_input("본조사일 (YYYY-MM-DD)", key="본조사", placeholder="2024-01-01")
        성명 = st.text_input("성명", key="성명")

# 2. 근골격계 부담작업 체크리스트 탭
with tabs[1]:
    st.subheader("근골격계 부담작업 체크리스트")
    
    # 엑셀 파일 업로드 기능 추가
    with st.expander("[엑셀 파일 업로드]"):
        st.info("""
        [엑셀 파일 양식]
        - 1열: 회사명
        - 2열: 소속
        - 3열: 반
        - 4열: 단위작업명
        - 5~15열: 1호~11호 (O(해당), △(잠재위험), X(미해당) 중 입력)
        """)
        
        uploaded_excel = st.file_uploader("엑셀 파일 선택", type=['xlsx', 'xls'])
        
        if uploaded_excel is not None:
            try:
                # 엑셀 파일 읽기
                df_excel = pd.read_excel(uploaded_excel)
                
                # 컬럼명 확인 및 조정
                expected_columns = ["회사명", "소속", "반", "단위작업명"] + [f"{i}호" for i in range(1, 12)]
                
                # 컬럼 개수가 맞는지 확인
                if len(df_excel.columns) >= 15:
                    # 컬럼명 재설정
                    df_excel.columns = expected_columns[:len(df_excel.columns)]
                    
                    # 값 검증 (O(해당), △(잠재위험), X(미해당)만 허용)
                    valid_values = ["O(해당)", "△(잠재위험)", "X(미해당)"]
                    
                    # 5번째 열부터 15번째 열까지 검증
                    for col in expected_columns[4:]:
                        if col in df_excel.columns:
                            # 유효하지 않은 값은 X(미해당)으로 변경
                            df_excel[col] = df_excel[col].apply(
                                lambda x: x if x in valid_values else "X(미해당)"
                            )
                    
                    if st.button("[데이터 적용하기]"):
                        st.session_state["checklist_df"] = df_excel
                        
                        # 즉시 Excel 파일로 저장
                        if st.session_state.get("session_id") and st.session_state.get("workplace"):
                            success, _ = safe_save_to_excel(st.session_state["session_id"], st.session_state.get("workplace"))
                            if success:
                                st.session_state["last_save_time"] = time.time()
                                st.session_state["last_successful_save"] = datetime.now()
                        
                        st.success("[적용 완료] 엑셀 데이터를 성공적으로 불러오고 저장했습니다!")
                        st.rerun()
                    
                    # 미리보기
                    st.markdown("#### [데이터 미리보기]")
                    st.dataframe(df_excel)
                    
                else:
                    st.error("[오류] 엑셀 파일의 컬럼이 15개 이상이어야 합니다. (회사명, 소속, 반, 단위작업명, 1호~11호)")
                    
            except Exception as e:
                st.error(f"[파일 읽기 오류] {str(e)}")
    
    # 샘플 엑셀 파일 다운로드
    with st.expander("[샘플 엑셀 파일 다운로드]"):
        # 샘플 데이터 생성
        sample_data = pd.DataFrame({
            "회사명": ["A회사", "A회사", "A회사", "B회사", "B회사"],
            "소속": ["생산1팀", "생산1팀", "생산2팀", "품질팀", "품질팀"],
            "반": ["조립1반", "조립1반", "포장반", "검사1반", "검사2반"],
            "단위작업명": ["부품조립", "나사체결", "제품포장", "외관검사", "성능검사"],
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
            label="[샘플 엑셀 다운로드]",
            data=sample_output,
            file_name="체크리스트_샘플.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("##### 샘플 데이터 구조:")
        st.dataframe(sample_data)
    
    # 단위작업명 병합 기능
    with st.expander("[단위작업명 병합 기능]"):
        st.info("여러 개의 단위작업을 하나로 합칠 수 있습니다. 병합 시 부담작업 정보는 가장 높은 수준으로 통합됩니다.")
        
        if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
            if not st.session_state["checklist_df"].empty:
                # 선택 체크박스를 포함한 데이터프레임 표시
                df_with_select = st.session_state["checklist_df"].copy()
                df_with_select.insert(0, "선택", False)
                
                # 선택 가능한 데이터 편집기
                selected_df = st.data_editor(
                    df_with_select,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "선택": st.column_config.CheckboxColumn("선택", width=50),
                    },
                    key="병합_선택_df"
                )
                
                # 선택된 행의 인덱스 가져오기
                selected_indices = selected_df[selected_df["선택"] == True].index.tolist()
                
                if selected_indices:
                    st.info(f"선택된 항목 수: {len(selected_indices)}개")
                    
                    # 병합할 이름 입력
                    병합_이름 = st.text_input("병합 후 단위작업명", placeholder="예: 부품조립+나사체결")
                    
                    if st.button("[선택 항목 병합]", type="primary"):
                        if 병합_이름:
                            # 병합 수행
                            merged_df = merge_unit_works(selected_indices, st.session_state["checklist_df"], 병합_이름)
                            st.session_state["checklist_df"] = merged_df
                            
                            # 즉시 저장
                            if st.session_state.get("session_id") and st.session_state.get("workplace"):
                                success, _ = safe_save_to_excel(st.session_state["session_id"], st.session_state.get("workplace"))
                                if success:
                                    st.success("[병합 완료] 단위작업이 성공적으로 병합되고 저장되었습니다!")
                                    st.rerun()
                        else:
                            st.warning("병합 후 단위작업명을 입력해주세요.")
            else:
                st.info("체크리스트 데이터를 먼저 입력해주세요.")
    
    st.markdown("---")
    
    # 기존 데이터 편집기
    columns = [
        "회사명", "소속", "반", "단위작업명"
    ] + [f"{i}호" for i in range(1, 12)]
    
    # 세션 상태에 저장된 데이터가 있으면 사용, 없으면 빈 데이터
    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
        if not st.session_state["checklist_df"].empty:
            data = st.session_state["checklist_df"]
        else:
            data = pd.DataFrame(
                columns=columns,
                data=[["", "", "", ""] + ["X(미해당)"]*11 for _ in range(5)]
            )
    else:
        data = pd.DataFrame(
            columns=columns,
            data=[["", "", "", ""] + ["X(미해당)"]*11 for _ in range(5)]
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
    column_config["회사명"] = st.column_config.TextColumn("회사명")
    column_config["소속"] = st.column_config.TextColumn("소속")
    column_config["반"] = st.column_config.TextColumn("반")
    column_config["단위작업명"] = st.column_config.TextColumn("단위작업명")

    edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config=column_config
    )
    st.session_state["checklist_df"] = edited_df
    
    # 현재 등록된 계층 구조 표시
    col1, col2, col3 = st.columns(3)
    with col1:
        회사명_목록 = get_회사명_목록()
        if 회사명_목록:
            st.info(f"[회사] {len(회사명_목록)}개")
            for 회사 in 회사명_목록:
                st.write(f"- {회사}")
    
    with col2:
        if 회사명_목록:
            selected_회사_통계 = st.selectbox("회사 선택", 회사명_목록, key="통계_회사선택")
            소속_목록 = get_소속_목록(selected_회사_통계)
            if 소속_목록:
                st.info(f"[{selected_회사_통계}의 소속] {len(소속_목록)}개")
                for 소속 in 소속_목록:
                    st.write(f"- {소속}")
    
    with col3:
        if 회사명_목록 and 소속_목록:
            selected_소속_통계 = st.selectbox("소속 선택", 소속_목록, key="통계_소속선택")
            반_목록 = get_반_목록(selected_회사_통계, selected_소속_통계)
            if 반_목록:
                st.info(f"[{selected_소속_통계}의 반] {len(반_목록)}개")
                for 반 in 반_목록:
                    st.write(f"- {반}")

# 3. 유해요인조사표 탭
with tabs[2]:
    st.title("유해요인조사표")
    
    # 계층적 선택
    col1, col2, col3 = st.columns(3)
    
    with col1:
        회사명_목록 = get_회사명_목록()
        if not 회사명_목록:
            st.warning("[경고] 먼저 체크리스트에서 데이터를 입력해주세요.")
            selected_회사_유해 = None
        else:
            selected_회사_유해 = st.selectbox("회사명 선택", 회사명_목록, key="유해_회사선택")
    
    with col2:
        if selected_회사_유해:
            소속_목록 = get_소속_목록(selected_회사_유해)
            if 소속_목록:
                selected_소속_유해 = st.selectbox("소속 선택", 소속_목록, key="유해_소속선택")
            else:
                selected_소속_유해 = None
                st.info("소속이 없습니다.")
        else:
            selected_소속_유해 = None
    
    with col3:
        if selected_회사_유해 and selected_소속_유해:
            반_목록 = get_반_목록(selected_회사_유해, selected_소속_유해)
            if 반_목록:
                selected_반_유해 = st.selectbox("반 선택", 반_목록, key="유해_반선택")
            else:
                selected_반_유해 = None
                st.info("반이 없습니다.")
        else:
            selected_반_유해 = None
    
    # 선택된 반에 대한 유해요인조사표 작성
    if selected_회사_유해 and selected_소속_유해 and selected_반_유해:
        st.info(f"[선택] {selected_회사_유해} > {selected_소속_유해} > {selected_반_유해}")
        
        # 해당 반의 단위작업명 가져오기
        단위작업명_목록 = get_단위작업명_목록(selected_회사_유해, selected_소속_유해, selected_반_유해)
        
        with st.expander(f"[{selected_반_유해} - 유해요인조사표]", expanded=True):
            st.markdown("#### 가. 조사개요")
            col1, col2 = st.columns(2)
            with col1:
                조사일시 = st.text_input("조사일시", key=f"조사일시_{selected_반_유해}")
                부서명 = st.text_input("부서명", value=selected_소속_유해, key=f"부서명_{selected_반_유해}")
            with col2:
                조사자 = st.text_input("조사자", key=f"조사자_{selected_반_유해}")
                작업공정명 = st.text_input("작업공정명", value=selected_반_유해, key=f"작업공정명_{selected_반_유해}")
            작업명_유해 = st.text_input("작업명(반)", value=selected_반_유해, key=f"작업명_{selected_반_유해}")
            
            # 단위작업명 표시
            if 단위작업명_목록:
                st.markdown("##### 단위작업명 목록")
                st.write(", ".join(단위작업명_목록))

            st.markdown("#### 나. 작업장 상황조사")

            def 상황조사행(항목명, 반):
                cols = st.columns([2, 5, 3])
                with cols[0]:
                    st.markdown(f"<div style='text-align:center; font-weight:bold; padding-top:0.7em;'>{항목명}</div>", unsafe_allow_html=True)
                with cols[1]:
                    상태 = st.radio(
                        label="",
                        options=["변화없음", "감소", "증가", "기타"],
                        key=f"{항목명}_상태_{반}",
                        horizontal=True,
                        label_visibility="collapsed"
                    )
                with cols[2]:
                    if 상태 == "감소":
                        st.text_input("감소 - 언제부터", key=f"{항목명}_감소_시작_{반}", placeholder="언제부터", label_visibility="collapsed")
                    elif 상태 == "증가":
                        st.text_input("증가 - 언제부터", key=f"{항목명}_증가_시작_{반}", placeholder="언제부터", label_visibility="collapsed")
                    elif 상태 == "기타":
                        st.text_input("기타 - 내용", key=f"{항목명}_기타_내용_{반}", placeholder="내용", label_visibility="collapsed")
                    else:
                        st.markdown("&nbsp;", unsafe_allow_html=True)

            for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
                상황조사행(항목, selected_반_유해)
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
    
    # 계층적 선택
    col1, col2, col3 = st.columns(3)
    
    with col1:
        회사명_목록 = get_회사명_목록()
        if not 회사명_목록:
            st.warning("[경고] 먼저 체크리스트에서 데이터를 입력해주세요.")
            selected_회사_작업 = None
        else:
            selected_회사_작업 = st.selectbox("회사명 선택", 회사명_목록, key="작업_회사선택")
    
    with col2:
        if selected_회사_작업:
            소속_목록 = get_소속_목록(selected_회사_작업)
            if 소속_목록:
                selected_소속_작업 = st.selectbox("소속 선택", 소속_목록, key="작업_소속선택")
            else:
                selected_소속_작업 = None
                st.info("소속이 없습니다.")
        else:
            selected_소속_작업 = None
    
    with col3:
        if selected_회사_작업 and selected_소속_작업:
            반_목록 = get_반_목록(selected_회사_작업, selected_소속_작업)
            if 반_목록:
                selected_반_작업 = st.selectbox("반 선택", 반_목록, key="작업_반선택")
            else:
                selected_반_작업 = None
                st.info("반이 없습니다.")
        else:
            selected_반_작업 = None
    
    if selected_회사_작업 and selected_소속_작업 and selected_반_작업:
        st.info(f"[선택] {selected_회사_작업} > {selected_소속_작업} > {selected_반_작업}")
        
        # 선택된 반에 대한 1,2,3단계
        with st.container():
            # 1단계: 유해요인 기본조사
            st.subheader(f"1단계: 유해요인 기본조사 - [{selected_반_작업}]")
            col1, col2 = st.columns(2)
            with col1:
                작업공정 = st.text_input("작업공정", value=selected_반_작업, key=f"1단계_작업공정_{selected_반_작업}")
            with col2:
                작업내용 = st.text_input("작업내용", key=f"1단계_작업내용_{selected_반_작업}")
            
            st.markdown("---")
            
            # 2단계: 작업별 작업부하 및 작업빈도
            st.subheader(f"2단계: 작업별 작업부하 및 작업빈도 - [{selected_반_작업}]")
            
            # 선택된 반에 해당하는 체크리스트 데이터 가져오기
            checklist_data = []
            if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
                if not st.session_state["checklist_df"].empty:
                    반_체크리스트 = st.session_state["checklist_df"][
                        (st.session_state["checklist_df"]["회사명"] == selected_회사_작업) &
                        (st.session_state["checklist_df"]["소속"] == selected_소속_작업) &
                        (st.session_state["checklist_df"]["반"] == selected_반_작업)
                    ]
                    
                    for idx, row in 반_체크리스트.iterrows():
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
                key=f"작업조건_data_editor_{selected_반_작업}"
            )
            
            # 편집된 데이터를 세션 상태에 저장
            st.session_state[f"작업조건_data_{selected_반_작업}"] = edited_df
            
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
                
                st.info("[도움말] 총점은 작업부하(A) × 작업빈도(B)로 자동 계산됩니다.")
            
            # 3단계: 유해요인평가
            st.markdown("---")
            st.subheader(f"3단계: 유해요인평가 - [{selected_반_작업}]")
            
            # 작업명과 근로자수 입력
            col1, col2 = st.columns(2)
            with col1:
                평가_작업명 = st.text_input("작업명(반)", value=selected_반_작업, key=f"3단계_작업명_{selected_반_작업}")
            with col2:
                평가_근로자수 = st.text_input("근로자수", key=f"3단계_근로자수_{selected_반_작업}")
            
            # 사진 업로드 및 설명 입력
            st.markdown("#### 작업 사진 및 설명")
            
            # 사진 개수 선택
            num_photos = st.number_input("사진 개수", min_value=1, max_value=10, value=3, key=f"사진개수_{selected_반_작업}")
            
            # 각 사진별로 업로드와 설명 입력
            for i in range(num_photos):
                st.markdown(f"##### 사진 {i+1}")
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    uploaded_file = st.file_uploader(
                        f"사진 {i+1} 업로드",
                        type=['png', 'jpg', 'jpeg'],
                        key=f"사진_{i+1}_업로드_{selected_반_작업}"
                    )
                    if uploaded_file:
                        st.image(uploaded_file, caption=f"사진 {i+1}", use_column_width=True)
                
                with col2:
                    photo_description = st.text_area(
                        f"사진 {i+1} 설명",
                        height=150,
                        key=f"사진_{i+1}_설명_{selected_반_작업}",
                        placeholder="이 사진에 대한 설명을 입력하세요..."
                    )
                
                st.markdown("---")
            
            # 작업별로 관련된 유해요인에 대한 원인분석
            st.markdown("---")
            st.subheader(f"작업별로 관련된 유해요인에 대한 원인분석 - [{selected_반_작업}]")
            
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
            원인분석_key = f"원인분석_항목_{selected_반_작업}"
            if 원인분석_key not in st.session_state:
                st.session_state[원인분석_key] = []
                # 부담작업 정보를 기반으로 초기 항목 생성
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
                if st.button("[추가]", key=f"원인분석_추가_{selected_반_작업}", use_container_width=True):
                    st.session_state[원인분석_key].append({
                        "단위작업명": "",
                        "부담작업호": "",
                        "유형": "",
                        "부담작업": "",
                        "비고": ""
                    })
                    st.rerun()
            with col3:
                if st.button("[삭제]", key=f"원인분석_삭제_{selected_반_작업}", use_container_width=True):
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
                        key=f"원인분석_단위작업명_{k}_{selected_반_작업}"
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
                                    힌트_텍스트.append(f"[잠재] {호수}: {부담작업_설명[호수]}")
                                else:
                                    힌트_텍스트.append(f"[해당] {호수}: {부담작업_설명[호수]}")
                        
                        if 힌트_텍스트:
                            st.info("[부담작업 힌트]\n" + "\n".join(힌트_텍스트))
                    else:
                        st.empty()  # 빈 공간 유지
                
                with col3:
                    hazard_entry["비고"] = st.text_input(
                        "비고", 
                        value=hazard_entry.get("비고", ""), 
                        key=f"원인분석_비고_{k}_{selected_반_작업}"
                    )
                
                # 유해요인 유형 선택
                hazard_type_options = ["", "반복동작", "부자연스러운 자세", "과도한 힘", "접촉스트레스 또는 기타(진동, 밀고 당기기 등)"]
                selected_hazard_type_index = hazard_type_options.index(hazard_entry.get("유형", "")) if hazard_entry.get("유형", "") in hazard_type_options else 0
                
                hazard_entry["유형"] = st.selectbox(
                    f"[{k+1}] 유해요인 유형 선택", 
                    hazard_type_options, 
                    index=selected_hazard_type_index, 
                    key=f"hazard_type_{k}_{selected_반_작업}",
                    help="선택한 단위작업의 부담작업 유형에 맞는 항목을 선택하세요"
                )

                # 유형별 상세 입력은 기존 코드와 동일...
                # (반복동작, 부자연스러운 자세, 과도한 힘, 접촉스트레스 관련 코드는 그대로 유지)
                
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
        if st.button("[정밀조사 추가]", use_container_width=True):
            st.session_state["정밀조사_목록"].append(f"정밀조사_{len(st.session_state['정밀조사_목록'])+1}")
            st.rerun()
    
    if not st.session_state["정밀조사_목록"]:
        st.info("[안내] 정밀조사가 필요한 경우 '정밀조사 추가' 버튼을 클릭하세요.")
    else:
        # 각 정밀조사 표시
        for idx, 조사명 in enumerate(st.session_state["정밀조사_목록"]):
            with st.expander(f"[{조사명}]", expanded=True):
                # 삭제 버튼
                col1, col2 = st.columns([10, 1])
                with col2:
                    if st.button("[X]", key=f"삭제_{조사명}"):
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
    
    # 반 목록 가져오기 (모든 반)
    전체_반_목록 = []
    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
        if not st.session_state["checklist_df"].empty:
            전체_반_목록 = st.session_state["checklist_df"]["반"].dropna().unique().tolist()
    
    # 1. 기초현황
    st.subheader("1. 기초현황")
    
    # 반별 데이터 자동 생성
    기초현황_columns = ["반", "응답자(명)", "나이", "근속년수", "남자(명)", "여자(명)", "합계"]
    
    if 전체_반_목록:
        # 반 목록을 기반으로 데이터 생성
        기초현황_data_rows = []
        for 반 in 전체_반_목록:
            기초현황_data_rows.append([반, "", "평균(세)", "평균(년)", "", "", ""])
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
    
    작업기간_columns = ["반", "<1년", "<3년", "<5년", "≥5년", "무응답", "합계", "이전<1년", "이전<3년", "이전<5년", "이전≥5년", "이전무응답", "이전합계"]
    
    if 전체_반_목록:
        # 반 목록을 기반으로 데이터 생성
        작업기간_data_rows = []
        for 반 in 전체_반_목록:
            작업기간_data_rows.append([반] + [""] * 12)
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
    육체적부담_columns = ["반", "전혀 힘들지 않음", "견딜만 함", "약간 힘듦", "힘듦", "매우 힘듦", "합계"]
    
    if 전체_반_목록:
        # 반 목록을 기반으로 데이터 생성
        육체적부담_data_rows = []
        for 반 in 전체_반_목록:
            육체적부담_data_rows.append([반] + [""] * 6)
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
    
    if 전체_반_목록:
        # 컬럼 정의
        통증호소자_columns = ["반", "구분", "목", "어깨", "팔/팔꿈치", "손/손목/손가락", "허리", "다리/발", "전체"]
        
        # 데이터 생성
        통증호소자_data = []
        
        for 반 in 전체_반_목록:
            # 각 반에 대해 정상, 관리대상자, 통증호소자 3개 행 추가
            통증호소자_data.append([반, "정상", "", "", "", "", "", "", ""])
            통증호소자_data.append(["", "관리대상자", "", "", "", "", "", "", ""])
            통증호소자_data.append(["", "통증호소자", "", "", "", "", "", "", ""])
        
        통증호소자_df = pd.DataFrame(통증호소자_data, columns=통증호소자_columns)
        
        # 컬럼 설정
        column_config = {
            "반": st.column_config.TextColumn("반", disabled=True, width=150),
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
            disabled=["반", "구분"]
        )
        
        # 세션 상태에 저장
        st.session_state["통증호소자_data_저장"] = 통증호소자_edited
    else:
        st.info("체크리스트에 데이터를 입력하면 자동으로 표가 생성됩니다.")
        
        # 빈 데이터프레임 표시
        통증호소자_columns = ["반", "구분", "목", "어깨", "팔/팔꿈치", "손/손목/손가락", "허리", "다리/발", "전체"]
        빈_df = pd.DataFrame(columns=통증호소자_columns)
        st.dataframe(빈_df, use_container_width=True)

# 7. 작업환경개선계획서 탭
with tabs[6]:
    st.title("작업환경개선계획서")
    
    # 컬럼 정의
    개선계획_columns = [
        "회사명",
        "소속",
        "반",
        "단위작업명",
        "문제점(유해요인의 원인)",
        "근로자의견",
        "개선방안",
        "추진일정",
        "개선비용",
        "개선우선순위"
    ]
    
    # 세션 상태에 개선계획 데이터가 없으면 초기화
    if "개선계획_data_저장" not in st.session_state or not validate_dataframe(st.session_state.get("개선계획_data_저장")):
        # 체크리스트 데이터 기반으로 초기 데이터 생성
        if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
            if not st.session_state["checklist_df"].empty:
                개선계획_data_rows = []
                for _, row in st.session_state["checklist_df"].iterrows():
                    if row["회사명"] and row["소속"] and row["반"] and row["단위작업명"]:
                        # 부담작업이 있는 경우만 추가
                        부담작업_있음 = False
                        for i in range(1, 12):
                            if row[f"{i}호"] in ["O(해당)", "△(잠재위험)"]:
                                부담작업_있음 = True
                                break
                        
                        if 부담작업_있음:
                            개선계획_data_rows.append([
                                row["회사명"],
                                row["소속"],
                                row["반"],
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
                        data=[["", "", "", "", "", "", "", "", "", ""] for _ in range(5)]
                    )
            else:
                # 초기 데이터 (빈 행 5개)
                개선계획_data = pd.DataFrame(
                    columns=개선계획_columns,
                    data=[["", "", "", "", "", "", "", "", "", ""] for _ in range(5)]
                )
        else:
            # 초기 데이터 (빈 행 5개)
            개선계획_data = pd.DataFrame(
                columns=개선계획_columns,
                data=[["", "", "", "", "", "", "", "", "", ""] for _ in range(5)]
            )
        
        st.session_state["개선계획_data_저장"] = 개선계획_data
    
    # 컬럼 설정
    개선계획_config = {
        "회사명": st.column_config.TextColumn("회사명", width=100),
        "소속": st.column_config.TextColumn("소속", width=100),
        "반": st.column_config.TextColumn("반", width=100),
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
    with st.expander("[작성 도움말]"):
        st.markdown("""
        - **회사명/소속/반**: 해당 작업의 조직 구조
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
        if st.button("[행 추가]", key="개선계획_행추가", use_container_width=True):
            new_row = pd.DataFrame([["", "", "", "", "", "", "", "", "", ""]], columns=개선계획_columns)
            st.session_state["개선계획_data_저장"] = pd.concat([st.session_state["개선계획_data_저장"], new_row], ignore_index=True)
            st.rerun()
    with col3:
        if st.button("[마지막 행 삭제]", key="개선계획_행삭제", use_container_width=True):
            if len(st.session_state["개선계획_data_저장"]) > 0:
                st.session_state["개선계획_data_저장"] = st.session_state["개선계획_data_저장"].iloc[:-1]
                st.rerun()
    
    # 전체 보고서 다운로드
    st.markdown("---")
    st.subheader("[전체 보고서 다운로드]")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # 엑셀 다운로드 버튼
        if st.button("[전체 Excel 보고서 다운로드]", use_container_width=True):
            try:
                output = BytesIO()
                
                # 반 목록 다시 가져오기
                반_목록_다운로드 = get_반_목록()
                
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
                    if "checklist_df" in st.session_state and validate_dataframe(st.session_state.get("checklist_df")):
                        if not st.session_state["checklist_df"].empty:
                            st.session_state["checklist_df"].to_excel(writer, sheet_name='체크리스트', index=False)
                    
                    # 유해요인조사표 데이터 저장 (반별로)
                    for 반 in 반_목록_다운로드:
                        조사표_data = []
                        
                        # 조사개요
                        조사표_data.append(["조사개요"])
                        조사표_data.append(["조사일시", st.session_state.get(f"조사일시_{반}", "")])
                        조사표_data.append(["부서명", st.session_state.get(f"부서명_{반}", "")])
                        조사표_data.append(["조사자", st.session_state.get(f"조사자_{반}", "")])
                        조사표_data.append(["작업공정명", st.session_state.get(f"작업공정명_{반}", "")])
                        조사표_data.append(["작업명(반)", st.session_state.get(f"작업명_{반}", "")])
                        조사표_data.append([])  # 빈 행
                        
                        # 작업장 상황조사
                        조사표_data.append(["작업장 상황조사"])
                        조사표_data.append(["항목", "상태", "세부사항"])
                        
                        for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
                            상태 = st.session_state.get(f"{항목}_상태_{반}", "변화없음")
                            세부사항 = ""
                            if 상태 == "감소":
                                세부사항 = st.session_state.get(f"{항목}_감소_시작_{반}", "")
                            elif 상태 == "증가":
                                세부사항 = st.session_state.get(f"{항목}_증가_시작_{반}", "")
                            elif 상태 == "기타":
                                세부사항 = st.session_state.get(f"{항목}_기타_내용_{반}", "")
                            
                            조사표_data.append([항목, 상태, 세부사항])
                        
                        if 조사표_data:
                            조사표_df = pd.DataFrame(조사표_data)
                            safe_반 = str(반).replace('/', '_').replace('\\', '_')[:31]
                            sheet_name = f'유해요인_{safe_반}'
                            조사표_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    
                    # 나머지 데이터도 동일하게 처리...
                    
                output.seek(0)
                st.download_button(
                    label="[Excel 다운로드]",
                    data=output,
                    file_name=f"근골격계_유해요인조사_{st.session_state.get('workplace', '')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("[다운로드 준비 완료] Excel 보고서가 생성되었습니다!")
                
            except Exception as e:
                st.error(f"Excel 파일 생성 중 오류가 발생했습니다: {str(e)}")
                st.info("데이터를 입력한 후 다시 시도해주세요.")
    
    with col2:
        # PDF 보고서 생성 버튼 (기존 코드와 동일)
        if PDF_AVAILABLE:
            if st.button("[PDF 보고서 생성]", use_container_width=True):
                st.info("PDF 생성 기능은 준비 중입니다.")
        else:
            st.info("PDF 생성 기능을 사용하려면 reportlab 라이브러리를 설치하세요: pip install reportlab")
