import streamlit as st
import pandas as pd
import json
import io
from typing import Dict, List, Any

st.set_page_config(
    page_title="Excel to JSON",
    layout="wide"
)

def extract_excel_metadata(file) -> Dict[str, Dict[str, List[str]]]:
    """엑셀 파일에서 메타데이터를 추출합니다.
    Returns: {sheet_name: [columns]}
    """
    metadata = {}
    
    # 모든 시트 읽기
    excel_file = pd.ExcelFile(file)
    
    for sheet_name in excel_file.sheet_names:
        # 첫 번째 행을 헤더로 읽기
        df_with_header = pd.read_excel(file, sheet_name=sheet_name)
        
        # 컬럼명 가져오기
        columns = [str(col) for col in df_with_header.columns.tolist()]
        
        metadata[sheet_name] = columns
    
    return metadata

def convert_to_json(all_files_metadata: Dict[str, Dict[str, Dict[str, List[str]]]]) -> str:
    """여러 파일의 메타데이터를 JSON 형식으로 변환합니다.
    Structure: {filename: {sheet_name: [columns]}}
    """
    return json.dumps(all_files_metadata, ensure_ascii=False, indent=2)

def main():
    st.title("Excel to JSON")
    st.markdown("여러 엑셀 파일의 구조 정보를 JSON 형식으로 변환합니다")
    
    # 사이드바 설정
    with st.sidebar:
        st.header("설정")
        
        st.markdown("---")
        st.markdown("### 사용 방법")
        st.markdown("""
        1. 여러 엑셀 파일을 업로드하세요
        2. 파일 구조를 확인하세요
        3. JSON으로 다운로드하세요
        
        **JSON 구조:**
        ```
        {
          "파일명": {
            "시트명": [컬럼1, 컬럼2, ...]
          }
        }
        ```
        """)
    
    # 메인 컨텐츠
    st.header("파일 업로드")
    uploaded_files = st.file_uploader(
        "엑셀 파일들을 선택하세요 (여러 개 선택 가능)",
        type=['xlsx', 'xls'],
        help="지원 형식: .xlsx, .xls",
        accept_multiple_files=True
    )
    
    if uploaded_files:
        try:
            all_files_metadata = {}
            
            # 각 파일 처리
            with st.spinner("파일들을 분석중입니다..."):
                for uploaded_file in uploaded_files:
                    # 파일명에서 확장자 제거
                    file_name = uploaded_file.name.rsplit('.', 1)[0]
                    metadata = extract_excel_metadata(uploaded_file)
                    all_files_metadata[file_name] = metadata
            
            # 파일 정보 요약
            st.header("업로드된 파일 정보")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("총 파일 수", len(uploaded_files))
            with col2:
                total_sheets = sum(len(file_data) for file_data in all_files_metadata.values())
                st.metric("총 시트 수", total_sheets)
            with col3:
                total_columns = sum(
                    len(columns) 
                    for file_data in all_files_metadata.values() 
                    for columns in file_data.values()
                )
                st.metric("총 컬럼 수", total_columns)
            
            # JSON 변환
            json_str = convert_to_json(all_files_metadata)
            
            # 다운로드 버튼
            st.markdown("---")
            col_download1, col_download2, col_download3 = st.columns([1, 2, 1])
            with col_download2:
                st.download_button(
                    label="JSON 파일 다운로드",
                    data=json_str,
                    file_name="excel_metadata_combined.json",
                    mime="application/json",
                    use_container_width=True
                )
            
            # JSON 미리보기
            st.header("JSON 변환 결과")
            with st.expander("JSON 미리보기", expanded=True):
                st.code(json_str, language='json')
                
        except Exception as e:
            st.error(f"파일 처리 중 오류가 발생했습니다: {str(e)}")
            st.info("올바른 엑셀 파일인지 확인해주세요.")
    
    else:
        # 안내 메시지
        st.info("엑셀 파일들을 업로드하여 시작하세요 (여러 개 선택 가능)")
        

if __name__ == "__main__":
    main()