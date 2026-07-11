import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.title("셔틀버스 수요조사 자동 정리 프로그램")

st.write("수요조사 결과 파일과 빈 양식 엑셀을 업로드하세요.")

# csv 파일도 업로드 가능하도록 수정
survey_file = st.file_uploader("1. 수요조사 파일 업로드 (엑셀 또는 CSV)", type=['xlsx', 'csv'])
template_file = st.file_uploader("2. 셔틀 양식 엑셀 파일 업로드", type=['xlsx'])

if survey_file and template_file:
    if st.button("데이터 정리 시작"):
        try:
            # 1. 수요조사 데이터 읽기 (파일 확장자에 따라 다르게 읽기)
            if survey_file.name.endswith('.csv'):
                survey_df = pd.read_csv(survey_file)
            else:
                survey_df = pd.read_excel(survey_file)
            
            # --- [핵심 수정 부분] 날짜에 상관없이 '교번', '출교', '귀교' 단어로 열을 찾습니다 ---
            id_col = [col for col in survey_df.columns if '교번' in col][0]
            dep_col = [col for col in survey_df.columns if '출교' in col][0]
            ret_col = [col for col in survey_df.columns if '귀교' in col][0]
            
            survey_dict = {}
            for idx, row in survey_df.iterrows():
                student_id = str(row[id_col]).strip()
                dep = str(row[dep_col]).strip()
                ret = str(row[ret_col]).strip()
                survey_dict[student_id] = (dep, ret)

            def map_val(v):
                v = str(v).strip().lower()
                if v == 'o': return '○'
                if v == 'x': return 'X'
                return v

            # 2. 양식 파일에 데이터 쓰기
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active

            for row_idx in range(3, ws.max_row + 1):
                student_id_cell = ws.cell(row=row_idx, column=1)
                student_id = str(student_id_cell.value).strip()
                
                if student_id == 'None' or student_id == '':
                    continue
                    
                if student_id in survey_dict:
                    dep, ret = survey_dict[student_id]
                    ws.cell(row=row_idx, column=5).value = map_val(dep)
                    ws.cell(row=row_idx, column=6).value = map_val(ret)
                else:
                    ws.cell(row=row_idx, column=5).value = '답변X'
                    ws.cell(row=row_idx, column=6).value = '답변X'

            # 3. 결과물을 메모리 버퍼에 저장
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("데이터 정리가 완료되었습니다. 아래 버튼을 눌러 다운로드하세요.")
            st.download_button(
                label="정리본 엑셀 다운로드",
                data=output,
                file_name="셔틀_정리본.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다. 파일 양식을 확인해주세요: {e}")
