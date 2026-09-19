import streamlit as st
import docx
import re
import pandas as pd
import io
import zipfile

# 페이지 기본 설정
st.set_page_config(page_title="해양 수질검사 GPS 좌표 변환기 (DMS & DM)", page_icon="🌊", layout="wide")

st.title("🌊 해양 수질검사 시료채취기록부 GPS 좌표 추출 & 변환기")
st.write("시료채취기록부(`.docx`)를 업로드하면 **도분초(DMS)** 및 **도분(DM)** 좌표를 십진수(Decimal)로 자동 변환하여 **CSV** 및 **구글 어스 레이어(KML)** 파일로 생성합니다.")

# DMS(도분초) & DM(도분) -> Decimal 변환 통합 함수
def coord_to_dd(coord_str):
    if not coord_str or not isinstance(coord_str, str):
        return None
    coord_str = coord_str.strip()
    
    # 1. 도분초(DMS) 패턴 검사 (예: 36°23'20.87", 36° 23' 20.87")
    dms_match = re.search(r'(\d+)[\s°]+(\d+)[\s\'\′]+([\d\.]+)[\"\″]*', coord_str)
    if dms_match:
        deg = float(dms_match.group(1))
        minute = float(dms_match.group(2))
        second = float(dms_match.group(3))
        dd = deg + (minute / 60.0) + (second / 3600.0)
        return round(dd, 6)
    
    # 2. 도분(DM) 패턴 검사 (예: 36° 23.3478', 36°23.3478, 36 23.3478)
    dm_match = re.search(r'(\d+)[\s°]+([\d\.]+)[\'\′]*', coord_str)
    if dm_match:
        deg = float(dm_match.group(1))
        minute = float(dm_match.group(2))
        dd = deg + (minute / 60.0)
        return round(dd, 6)
        
    return None

# KML 문자열 생성 함수
def create_kml(data, task_name):
    kml_str = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>{task_name} 시료채취 조사지점</name>
"""
    for item in data:
        kml_str += f"""    <Placemark>
      <name>{item['조사지점']}</name>
      <description>과업명: {task_name} | 원본좌표: {item['위도(원본)']}, {item['경도(원본)']}</description>
      <Point>
        <coordinates>{item['경도(Decimal)']},{item['위도(Decimal)']},0</coordinates>
      </Point>
    </Placemark>
"""
    kml_str += """  </Document>
</kml>"""
    return kml_str

# 파일 업로더
uploaded_file = st.file_uploader("시료채취기록부 워드 문서(.docx)를 선택하세요", type=["docx"])

if uploaded_file is not None:
    try:
        doc = docx.Document(uploaded_file)
        task_data = {}
        current_task = "미지정지역"

        for table in doc.tables:
            table_text = ""
            for row in table.rows:
                for cell in row.cells:
                    table_text += cell.text + " "
            
            # 과업명 파싱
            if "과업명" in table_text:
                for row in table.rows:
                    for i, cell in enumerate(row.cells):
                        if "과업명" in cell.text and i + 1 < len(row.cells):
                            task_name_raw = row.cells[i+1].text.strip()
                            if task_name_raw:
                                match = re.search(r'\((.*?)\)', task_name_raw)
                                current_task = match.group(1) if match else task_name_raw
                                if current_task not in task_data:
                                    task_data[current_task] = []

            # 좌표 추출
            if "위도(N)" in table_text or "위도" in table_text:
                for row in table.rows:
                    row_cells = [c.text.strip() for c in row.cells]
                    unique_cells = []
                    for c in row_cells:
                        if not unique_cells or c != unique_cells[-1]:
                            unique_cells.append(c)
                    
                    lat, lon, station = None, None, None
                    for idx, text in enumerate(unique_cells):
                        if re.search(r'\d+°', text) or re.search(r'\d+\s+\d+\.', text):
                            if lat is None:
                                lat = text
                                if idx > 0:
                                    station = unique_cells[idx-1]
                            elif lon is None:
                                lon = text
                    
                    if station and lat and lon and station not in ['위치(좌표)', '조사지점', '위도(N)', '경도(E)', '경도(S)', '경도(N)']:
                        station_clean = station.replace('\n', ' ').strip()
                        lat_dd = coord_to_dd(lat)
                        lon_dd = coord_to_dd(lon)
                        
                        if current_task in task_data:
                            exists = any(d['조사지점'] == station_clean for d in task_data[current_task])
                            if not exists and lat_dd and lon_dd:
                                task_data[current_task].append({
                                    '과업명': current_task,
                                    '조사지점': station_clean,
                                    '위도(원본)': lat.strip(),
                                    '경도(원본)': lon.strip(),
                                    '위도(Decimal)': lat_dd,
                                    '경도(Decimal)': lon_dd
                                })

        # 결과 출력 및 다운로드
        if task_data:
            st.success("🎉 GPS 좌표(DMS/DM) 파싱 및 십진수 변환, CSV & KML 파일 생성이 완성되었습니다!")
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for task, data in task_data.items():
                    if data:
                        df = pd.DataFrame(data)
                        st.subheader(f"📍 지역: {task}")
                        st.dataframe(df, use_container_width=True)
                        
                        col1, col2 = st.columns(2)
                        
                        # CSV 버튼
                        csv_bytes = df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                        col1.download_button(
                            label=f"📥 {task} CSV 다운로드",
                            data=csv_bytes,
                            file_name=f"GPS_좌표_{task}.csv",
                            mime="text/csv"
                        )
                        
                        # KML 구글어스 버튼
                        kml_str = create_kml(data, task)
                        col2.download_button(
                            label=f"🌍 {task} 구글 어스(KML) 레이어 다운로드",
                            data=kml_str.encode('utf-8'),
                            file_name=f"GoogleEarth_{task}.kml",
                            mime="application/vnd.google-earth.kml+xml"
                        )
                        
                        zip_file.writestr(f"GPS_좌표_{task}.csv", csv_bytes)
                        zip_file.writestr(f"GoogleEarth_{task}.kml", kml_str.encode('utf-8'))

            st.divider()
            st.download_button(
                label="📦 모든 지역 CSV & KML 전체 파일 한 번에 다운로드 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="GPS_좌표_및_KML_전체.zip",
                mime="application/zip"
            )
        else:
            st.warning("문서에서 좌표 데이터를 찾지 못했습니다.")

    except Exception as e:
        st.error(f"오류가 발생했습니다: {str(e)}")
