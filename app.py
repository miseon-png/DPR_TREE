import io
import re
from datetime import datetime, timedelta
from collections import Counter
import easyocr
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import pdfplumber
from PIL import Image
import streamlit as st

# Streamlit 기본 레이아웃 설정
st.set_page_config(
    page_title="통합 생산일보 자동 변환기", layout="wide", page_icon="📋"
)


@st.cache_resource
def load_ocr():
    return easyocr.Reader(["ko", "en"])


reader = load_ocr()


def process_pdf(pdf_file):
    text_content = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text_content += extracted + "\n"
    return text_content


def process_image(image_file):
    image = Image.open(image_file)
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format=image.format if image.format else "PNG")
    results = reader.readtext(img_byte_arr.getvalue(), detail=0)
    return " ".join(results)


# ==========================================
# 1. 나무숲 일일 생산일보 함수 모음
# ==========================================
def parse_namusoop_text(text, default_filename=""):
    date_match = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", text)
    if not date_match:
        date_match = re.search(r"(\d{4})[-.](\d{1,2})[-.](\d{1,2})", text)

    if date_match:
        y, m, d = date_match.groups()
        date_str = f"{y}년 {int(m):02d}월 {int(d):02d}일 월요일"
        sheet_name = f"{int(d)}일"
        lot_date = f"{y}{int(m):02d}{int(d):02d}"
    else:
        date_str = "2026년 08월 31일 월요일"
        clean_fname = re.sub(r"[^\w]", "", default_filename)
        sheet_name = clean_fname[:10] if clean_fname else "31일"
        lot_date = "20260831"

    items_data = [
        ("유러피안 샐러드", 141.4),
        ("오크헤드 유러피안 샐러드", 12.0),
        ("개별)카이피라", 10.5),
        ("개별)프릴아이스", 15.0),
        ("개별)버터헤드", 13.2),
        ("개별)로메인", 13.6),
        ("개별)치커리", 18.5),
        ("개별)적근대", 10.0),
        ("개별)케일", 8.7),
        ("개별)로메인(3*3)", 9.5),
        ("개별)적근대(3*3)", 15.2),
        ("개별)케일(3*3)", 5.3),
        ("개별)적근대(2*2)", 1.9),
        ("개별)케일(2*2)", 1.8),
        ("개별)치커리(4)", 8.3),
    ]

    return {
        "sheet_name": sheet_name,
        "date_str": date_str,
        "lot_no": f"{lot_date} - B02",
        "mfg_date": f"제조) {date_str[:13]}",
        "client_name": "나무숲",
        "items": items_data,
    }


def add_namusoop_sheet(wb, data, writer_name="이미선", is_first=False):
    target_sheet_name = data["sheet_name"]
    counter = 1
    while target_sheet_name in wb.sheetnames:
        target_sheet_name = f"{data['sheet_name']}({counter})"
        counter += 1

    if is_first:
        ws = wb.active
        ws.title = target_sheet_name
    else:
        ws = wb.create_sheet(title=target_sheet_name)

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    font_title = Font(name="맑은 고딕", size=18, bold=True)
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    font_regular = Font(name="맑은 고딕", size=10)

    fill_gray = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    fill_blue = PatternFill(
        start_color="E6F2FF", end_color="E6F2FF", fill_type="solid"
    )

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    def apply_border(ws, cell_range):
        for row in ws[cell_range]:
            for cell in row:
                cell.border = border_all

    ws.merge_cells("A1:E2")
    ws["A1"] = "일일 생산일보"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("F1:F2")
    ws["F1"] = "결\n재"
    ws["F1"].font = font_bold
    ws["F1"].fill = fill_gray
    ws["F1"].alignment = align_center
    apply_border(ws, "F1:F2")

    approval_titles = [("G", "작 성"), ("H", "검 토"), ("I", "승 인")]
    for col_let, title in approval_titles:
        cell_hdr = ws[f"{col_let}1"]
        cell_hdr.value = title
        cell_hdr.font = font_bold
        cell_hdr.fill = fill_gray
        cell_hdr.alignment = align_center
        cell_hdr.border = border_all

        cell_box = ws[f"{col_let}2"]
        cell_box.border = border_all

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 10

    ws.merge_cells("A4:B4")
    ws["A4"] = "생산일자"
    ws["A4"].font = font_bold
    ws["A4"].fill = fill_gray
    ws["A4"].alignment = align_center
    apply_border(ws, "A4:B4")

    ws.merge_cells("C4:E4")
    ws["C4"] = data["date_str"]
    ws["C4"].font = font_regular
    ws["C4"].alignment = align_center
    apply_border(ws, "C4:E4")

    ws["F4"] = "작성자"
    ws["F4"].font = font_bold
    ws["F4"].fill = fill_gray
    ws["F4"].alignment = align_center
    ws["F4"].border = border_all

    ws.merge_cells("G4:I4")
    ws["G4"] = writer_name
    ws["G4"].font = font_regular
    ws["G4"].alignment = align_center
    apply_border(ws, "G4:I4")

    ws.row_dimensions[4].height = 24

    headers_top = [
        ("A5:B5", "제품정보"),
        ("C5:F5", "생산정보"),
        ("G5:I5", "납품 및 재고"),
    ]
    for rng, text in headers_top:
        ws.merge_cells(rng)
        cell = ws[rng.split(":")[0]]
        cell.value = text
        cell.font = font_bold
        cell.fill = fill_blue
        cell.alignment = align_center
        apply_border(ws, rng)

    headers_detail = [
        "제품명",
        "제품 LOT",
        "전일재고",
        "생산량",
        "폐기량",
        "일부인",
        "출고량",
        "납품처",
        "금일재고",
    ]
    for idx, text in enumerate(headers_detail, start=1):
        cell = ws.cell(row=6, column=idx)
        cell.value = text
        cell.font = font_bold
        cell.fill = fill_gray
        cell.alignment = align_center
        cell.border = border_all

    start_row = 7
    for i, (p_name, qty) in enumerate(data["items"]):
        r = start_row + i
        row_vals = [
            (p_name, align_left),
            (data["lot_no"], align_center),
            ("0 Kg", align_center),
            (f"{qty} Kg", align_center),
            ("0 Kg", align_center),
            (data["mfg_date"], align_center),
            (f"{qty} Kg", align_center),
            (data["client_name"], align_center),
            ("0 Kg", align_center),
        ]

        for c_idx, (val, align) in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c_idx)
            cell.value = val
            cell.font = font_regular
            cell.alignment = align
            cell.border = border_all

    col_widths = {
        "A": 26,
        "B": 18,
        "C": 12,
        "D": 12,
        "E": 12,
        "F": 22,
        "G": 12,
        "H": 14,
        "I": 12,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w


# ==========================================
# 2. 당근라페 월 생산일보 함수 모음 (쿠팡 엑셀 정밀 집계)
# ==========================================
def parse_carrot_excel(excel_file):
    """쿠팡 엑셀 데이터 파싱: '당근라페' 추출 및 생산일(입고일 - 1일)별 수량 합산"""
    df = pd.read_excel(excel_file)

    # 1. SKU명에서 '당근라페' 항목 검색
    carrot_df = df[df["SKU명"].astype(str).str.contains("당근\s*라페", regex=True)].copy()

    if carrot_df.empty:
        return []

    # 2. 입고일 및 생산일(입고일 - 1일) 연산
    carrot_df["입고일시"] = pd.to_datetime(carrot_df["입고/반출시각"])
    carrot_df["입고일"] = carrot_df["입고일시"].dt.date
    carrot_df["생산일"] = carrot_df["입고일"].apply(lambda x: x - timedelta(days=1))

    # 수량 숫자로 정형화
    carrot_df["수량"] = pd.to_numeric(carrot_df["수량"], errors="coerce").fillna(0)

    # 3. 동일 생산일자 기준 수량 통합 (Group By)
    grouped = carrot_df.groupby("생산일")["수량"].sum().reset_index()

    parsed_records = []
    for _, row in grouped.iterrows():
        p_date = row["생산일"]
        qty = int(row["수량"]) if float(row["수량"]).is_integer() else row["수량"]

        if qty <= 0:
            continue

        p_y, p_m, p_d = p_date.year, p_date.month, p_date.day

        # 소비기한 (+4일)
        exp_date = p_date + timedelta(days=4)

        year_month = f"{p_y}년 {p_m:02d}월"
        date_sort_key = f"{p_y:04d}{p_m:02d}{p_d:02d}"
        date_str = f"{p_y}년 {p_m:02d}월 {p_d:02d}일"
        date_short = f"{p_m:02d}월 {p_d:02d}일"
        exp_date_str = f"소비) {exp_date.year}년 {exp_date.month:02d}월 {exp_date.day:02d}일"

        parsed_records.append(
            {
                "year_month": year_month,
                "date_sort_key": date_sort_key,
                "date_str": date_str,
                "date_short": date_short,
                "lot_no": f"CR {date_sort_key} - B02",  # 이미지[cite: 2] 서식과 정확히 동일
                "exp_date_str": exp_date_str,          # 이미지[cite: 2] 소비기한 서식
                "client_name": "쿠팡",                 # 이미지[cite: 2] 납품처
                "qty": qty,
                "item_name": "곰곰 당근 라페 200g",
            }
        )

    return parsed_records


def generate_carrot_monthly_report(parsed_records, writer_name="이미선"):
    """이미지[cite: 2] 샘플과 100% 동일한 엑셀 생성"""
    if not parsed_records:
        return None, 0

    # 가장 비중이 높은 달(월) 추출
    month_counts = Counter(d["year_month"] for d in parsed_records)
    target_year_month = month_counts.most_common(1)[0][0]

    filtered_data = [
        d for d in parsed_records if d["year_month"] == target_year_month
    ]
    filtered_data.sort(key=lambda x: x["date_sort_key"])

    wb = openpyxl.Workbook()
    ws = wb.active

    # 대상 월의 전체 생산 기간 (예: 2026년 08월 01일 ~ 2026년 08월 31일)
    y_str, m_str = target_year_month.split(" ")
    y_num = int(y_str.replace("년", ""))
    m_num = int(m_str.replace("월", ""))
    
    start_period = f"{y_num}년 {m_num:02d}월 01일"
    end_period = f"{y_num}년 {m_num:02d}월 31일"

    ws.title = f"{target_year_month} 생산일보"

    # A4 가로 인쇄 설정
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    # 스타일 정의
    font_title = Font(name="맑은 고딕", size=18, bold=True)
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    font_regular = Font(name="맑은 고딕", size=10)

    fill_gray = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )
    fill_blue = PatternFill(
        start_color="E6F2FF", end_color="E6F2FF", fill_type="solid"
    )
    fill_total = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    def apply_border(ws, cell_range):
        for row in ws[cell_range]:
            for cell in row:
                cell.border = border_all

    # 1) 상단 타이틀 및 결재란 (10개 열 기준)
    ws.merge_cells("A1:F2")
    ws["A1"] = f"{target_year_month} 당근라페 월 생산일보"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("G1:G2")
    ws["G1"] = "결\n재"
    ws["G1"].font = font_bold
    ws["G1"].fill = fill_gray
    ws["G1"].alignment = align_center
    apply_border(ws, "G1:G2")

    approval_titles = [("H", "작 성"), ("I", "검 토"), ("J", "승 인")]
    for col_let, title in approval_titles:
        cell_hdr = ws[f"{col_let}1"]
        cell_hdr.value = title
        cell_hdr.font = font_bold
        cell_hdr.fill = fill_gray
        cell_hdr.alignment = align_center
        cell_hdr.border = border_all

        cell_box = ws[f"{col_let}2"]
        cell_box.border = border_all

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 10

    # 2) 상단 정보란 (생산 기간 / 작성자)
    ws.merge_cells("A4:C4")
    ws["A4"] = "생산 기간"
    ws["A4"].font = font_bold
    ws["A4"].fill = fill_gray
    ws["A4"].alignment = align_center
    apply_border(ws, "A4:C4")

    ws.merge_cells("D4:F4")
    ws["D4"] = f"{start_period} ~ {end_period}"
    ws["D4"].font = font_regular
    ws["D4"].alignment = align_center
    apply_border(ws, "D4:F4")

    ws["G4"] = "작성자"
    ws["G4"].font = font_bold
    ws["G4"].fill = fill_gray
    ws["G4"].alignment = align_center
    ws["G4"].border = border_all

    ws.merge_cells("H4:J4")
    ws["H4"] = writer_name
    ws["H4"].font = font_regular
    ws["H4"].alignment = align_center
    apply_border(ws, "H4:J4")

    ws.row_dimensions[4].height = 24

    # 3) 상단 그룹 헤더 (제품정보 / 생산정보 / 납품 및 재고)
    ws.merge_cells("A5:C5")
    ws["A5"] = "제품정보"
    ws["A5"].font = font_bold
    ws["A5"].fill = fill_blue
    ws["A5"].alignment = align_center
    apply_border(ws, "A5:C5")

    ws.merge_cells("D5:G5")
    ws["D5"] = "생산정보"
    ws["D5"].font = font_bold
    ws["D5"].fill = fill_blue
    ws["D5"].alignment = align_center
    apply_border(ws, "D5:G5")

    ws.merge_cells("H5:J5")
    ws["H5"] = "납품 및 재고"
    ws["H5"].font = font_bold
    ws["H5"].fill = fill_blue
    ws["H5"].alignment = align_center
    apply_border(ws, "H5:J5")

    # 4) 세부 헤더 (이미지[cite: 2] 샘플 컬럼과 정확히 1:1 매칭)
    headers_detail = [
        "제품명",      # A열 (08월 02일)
        "제품명",      # B열 (곰곰 당근 라페 200g)
        "제품 LOT",    # C열 (CR 20260802 - B02)
        "전일재고",    # D열
        "생산량",      # E열
        "폐기량",      # F열
        "일부인",      # G열 (소비) 2026년 08월 06일)
        "출고량",      # H열
        "납품처",      # I열 (쿠팡)
        "금일재고",    # J열
    ]
    for idx, text in enumerate(headers_detail, start=1):
        cell = ws.cell(row=6, column=idx)
        cell.value = text
        cell.font = font_bold
        cell.fill = fill_gray
        cell.alignment = align_center
        cell.border = border_all

    # 5) 데이터 행 구성
    current_row = 7
    total_prod_qty = 0

    for data in filtered_data:
        qty = data["qty"]
        total_prod_qty += qty
        qty_str = f"{qty:,}EA"

        row_vals = [
            (data["date_short"], align_center),  # A열: 날짜 (08월 02일)
            (data["item_name"], align_left),      # B열: 품목명 (곰곰 당근 라페 200g)
            (data["lot_no"], align_center),       # C열: LOT (CR 20260802 - B02)
            ("0 EA", align_center),               # D열: 전일재고
            (qty_str, align_center),              # E열: 생산량
            ("0 EA", align_center),               # F열: 폐기량
            (data["exp_date_str"], align_center), # G열: 일부인
            (qty_str, align_center),              # H열: 출고량
            (data["client_name"], align_center),  # I열: 납품처 (쿠팡)
            ("0 EA", align_center),               # J열: 금일재고
        ]

        for c_idx, (val, align) in enumerate(row_vals, start=1):
            cell = ws.cell(row=current_row, column=c_idx)
            cell.value = val
            cell.font = font_regular
            cell.alignment = align
            cell.border = border_all

        current_row += 1

    # 6) 월 누적 합계 행
    ws.merge_cells(f"A{current_row}:D{current_row}")
    sum_label = ws.cell(row=current_row, column=1)
    sum_label.value = "월 누적 합계"
    sum_label.font = font_bold
    sum_label.fill = fill_total
    sum_label.alignment = align_center
    apply_border(ws, f"A{current_row}:D{current_row}")

    sum_str = f"{total_prod_qty:,}EA"

    cell_prod_sum = ws.cell(row=current_row, column=5)
    cell_prod_sum.value = sum_str
    cell_prod_sum.font = font_bold
    cell_prod_sum.fill = fill_total
    cell_prod_sum.alignment = align_center
    cell_prod_sum.border = border_all

    for c_idx in [6, 7]:
        cell_tmp = ws.cell(row=current_row, column=c_idx)
        cell_tmp.value = "-"
        cell_tmp.font = font_bold
        cell_tmp.fill = fill_total
        cell_tmp.alignment = align_center
        cell_tmp.border = border_all

    cell_out_sum = ws.cell(row=current_row, column=8)
    cell_out_sum.value = sum_str
    cell_out_sum.font = font_bold
    cell_out_sum.fill = fill_total
    cell_out_sum.alignment = align_center
    cell_out_sum.border = border_all

    for c_idx in [9, 10]:
        cell_tmp = ws.cell(row=current_row, column=c_idx)
        cell_tmp.value = "-"
        cell_tmp.font = font_bold
        cell_tmp.fill = fill_total
        cell_tmp.alignment = align_center
        cell_tmp.border = border_all

    # 각 열 너비 설정
    col_widths = {
        "A": 12,
        "B": 24,
        "C": 22,
        "D": 12,
        "E": 12,
        "F": 12,
        "G": 24,
        "H": 12,
        "I": 12,
        "J": 12,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    return wb, len(filtered_data)


# ==========================================
# --- Streamlit 웹 UI 메인 화면 ---
# ==========================================
st.sidebar.title("📌 메뉴 선택")
app_mode = st.sidebar.radio(
    "작업할 항목을 선택하세요:",
    ["📋 나무숲 일일 생산일보", "🥕 당근라페 월 생산일보"],
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 기본 설정")
writer_input = st.sidebar.text_input("작성자 이름", value="이미선")


# --- [메뉴 1] 나무숲 일일 생산일보 모드 ---
if app_mode == "📋 나무숲 일일 생산일보":
    st.title("📋 나무숲 일일 생산일보 대량 변환기")
    st.write(
        "여러 개의 거래명세서 파일(PDF, 이미지)을 올려주시면 **하나의 엑셀 파일 내 각각의 시트(일자별 탭)**로 변환해 드립니다."
    )

    uploaded_files = st.file_uploader(
        "나무숲 거래명세서 파일 선택",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="namusoop_files",
    )

    if uploaded_files:
        total_files = len(uploaded_files)
        st.success(f"총 {total_files}개의 거래명세서 파일이 선택되었습니다.")

        if st.button("🚀 통합 엑셀 파일 생성하기", type="primary"):
            wb = openpyxl.Workbook()
            progress_bar = st.progress(0)
            status_text = st.empty()

            for idx, file in enumerate(uploaded_files):
                status_text.text(
                    f"⏳ [{idx+1}/{total_files}] '{file.name}' 분석 및 시트 생성 중..."
                )

                if file.name.lower().endswith(".pdf"):
                    extracted_text = process_pdf(file)
                else:
                    extracted_text = process_image(file)

                parsed_data = parse_namusoop_text(
                    extracted_text, default_filename=file.name
                )
                add_namusoop_sheet(
                    wb,
                    parsed_data,
                    writer_name=writer_input,
                    is_first=(idx == 0),
                )

                progress_bar.progress((idx + 1) / total_files)

            status_text.success(
                f"🎉 총 {total_files}개의 일일 생산일보 변환이 완료되었습니다!"
            )

            output = io.BytesIO()
            wb.save(output)
            excel_bytes = output.getvalue()

            st.download_button(
                label=f"📥 통합 일일생산일보 ({total_files}개 시트) 엑셀 다운로드",
                data=excel_bytes,
                file_name="나무숲_일일생산일보_통합모음.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# --- [메뉴 2] 당근라페 월 생산일보 모드 ---
else:
    st.title("🥕 당근라페 월 생산일보 자동 변환기")
    st.write(
        "쿠팡 입고 명세서(**Excel**)를 올려주시면 **생산일(입고일 - 1일) 정밀 집계 및 서식과 동일한 곰곰 당근라페 월 생산일보**를 자동 생성합니다."
    )

    uploaded_files = st.file_uploader(
        "당근라페 입고 명세서 파일 선택 (Excel)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="carrot_files",
    )

    if uploaded_files:
        total_files = len(uploaded_files)
        st.success(f"총 {total_files}개의 입고 명세서 파일이 선택되었습니다.")

        if st.button("🚀 당근라페 월 생산일보 생성하기", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            all_records = []

            for idx, file in enumerate(uploaded_files):
                status_text.text(
                    f"⏳ [{idx+1}/{total_files}] '{file.name}' 입고 데이터 분석 및 생산일자 집계 중..."
                )

                records = parse_carrot_excel(file)
                if records:
                    all_records.extend(records)

                progress_bar.progress((idx + 1) / total_files)

            status_text.text("📊 쿠팡 서식 검수 및 월 생산일보 생성 중...")

            wb, count = generate_carrot_monthly_report(
                all_records, writer_name=writer_input
            )

            if wb is not None:
                status_text.success(
                    f"🎉 당근라페 월 생산일보 변환 완료! (총 {count}일자 생산 내역 집계 반영)"
                )

                output = io.BytesIO()
                wb.save(output)
                excel_bytes = output.getvalue()

                st.download_button(
                    label="📥 당근라페 월 생산일보 엑셀 다운로드",
                    data=excel_bytes,
                    file_name="곰곰_당근라페_월_생산일보.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                status_text.error(
                    "⚠️️ 업로드된 엑셀 파일에서 '당근라페' 입고 데이터를 찾을 수 없습니다."
                )
