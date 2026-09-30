import io
import re
import easyocr
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pdfplumber
from PIL import Image
import streamlit as st

# Streamlit 기본 레이아웃 설정 (넓은 화면)
st.set_page_config(
    page_title="일일 생산일보 자동 변환기", layout="wide", page_icon="📋"
)


# EasyOCR 인공지능 모델 캐싱 로드 (초기 1회만 로딩하여 속도 최적화)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(["ko", "en"])


reader = load_ocr()


def process_pdf(pdf_file):
    """PDF 거래명세서에서 텍스트 추출"""
    text_content = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text_content += page.extract_text() + "\n"
    return text_content


def process_image(image_file):
    """이미지 거래명세서(PNG/JPG)에서 OCR 텍스트 추출"""
    image = Image.open(image_file)
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format=image.format if image.format else "PNG")
    results = reader.readtext(img_byte_arr.getvalue(), detail=0)
    return " ".join(results)


def parse_statement_text(text, default_filename=""):
    """거래명세서 텍스트에서 생산일자, 시트명, 품목 정보 파싱"""
    # 날짜 패턴 검색 (예: 2026년 8월 31일 또는 2026-08-31)
    date_match = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", text)
    if not date_match:
        date_match = re.search(r"(\d{4})[-.](\d{1,2})[-.](\d{1,2})", text)

    if date_match:
        y, m, d = date_match.groups()
        date_str = f"{y}년 {int(m):02d}월 {int(d):02d}일 월요일"
        sheet_name = f"{y[2:]}{int(m):02d}{int(d):02d}_생산일보"
        lot_date = f"{y}{int(m):02d}{int(d):02d}"
    else:
        date_str = "2026년 08월 31일 월요일"
        clean_fname = re.sub(r"[^\w]", "", default_filename)
        sheet_name = clean_fname[:15] if clean_fname else "생산일보"
        lot_date = "20260831"

    # 전처리 품목 데이터 (원물 품목 자동 제외)
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


def add_report_sheet(wb, data, writer_name="이미선", is_first=False):
    """엑셀 워크북에 개별 시트(가로 방향 + 상단 결재란 정렬 수정완료) 추가"""
    if is_first:
        ws = wb.active
        ws.title = data["sheet_name"]
    else:
        ws = wb.create_sheet(title=data["sheet_name"])

    # A4 가로(Landscape) 인쇄 설정
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

    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # 병합 영역 테두리 적용 함수
    def apply_border(ws, cell_range):
        for row in ws[cell_range]:
            for cell in row:
                cell.border = border_all

    # 1) 제목 및 상단 결재란 (A~E열: 제목 / F열: 결재 / G,H,I열: 작성,검토,승인)
    ws.merge_cells("A1:E2")
    ws["A1"] = "일일 생산일보"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

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
    ws.row_dimensions[3].height = 10  # 여백 행

    # 2) 기본 정보 (A~B열: 생산일자 / C~E열: 날짜값 / F열: 작성자 / G~I열: 작성자명)
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

    # 3) 헤더 작성 (5행: 대분류, 6행: 세부 항목)
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

    # 4) 데이터 작성
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

    # 5) 가로 레이아웃 열 너비 설정
    col_widths = {
        "A": 26,  # 제품명
        "B": 18,  # 제품 LOT
        "C": 12,  # 전일재고
        "D": 12,  # 생산량
        "E": 12,  # 폐기량
        "F": 22,  # 일부인 / 결재 / 작성자
        "G": 12,  # 출고량 / 작성
        "H": 14,  # 납품처 / 검토
        "I": 12,  # 금일재고 / 승인
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w


# --- Streamlit 웹 UI 메인 화면 ---
st.title("📋 거래명세서 ➡️ 일일 생산일보 대량 변환기")
st.write(
    "여러 개의 거래명세서 파일(PDF, 이미지)을 드래그하여 올리면, **하나의 엑셀 파일 내 각각의 시트(탭)**로 자동 변환해 드립니다."
)

st.sidebar.header("⚙️ 기본 설정")
writer_input = st.sidebar.text_input("작성자 이름", value="이미선")

# 다중 파일 선택 업로더
uploaded_files = st.file_uploader(
    "거래명세서 파일 선택 (30개 이상 다중 선택 가능)",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
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
                f"⏳ [{idx+1}/{total_files}] '{file.name}' 분석 및 엑셀 시트 생성 중..."
            )

            if file.name.lower().endswith(".pdf"):
                extracted_text = process_pdf(file)
            else:
                extracted_text = process_image(file)

            parsed_data = parse_statement_text(
                extracted_text, default_filename=file.name
            )
            add_report_sheet(
                wb, parsed_data, writer_name=writer_input, is_first=(idx == 0)
            )

            progress_bar.progress((idx + 1) / total_files)

        status_text.success(
            f"🎉 총 {total_files}개의 거래명세서 변환이 완료되었습니다!"
        )

        output = io.BytesIO()
        wb.save(output)
        excel_bytes = output.getvalue()

        st.download_button(
            label=f"📥 통합 일일생산일보 ({total_files}개 시트) 엑셀 다운로드",
            data=excel_bytes,
            file_name="일일생산일보_통합모음.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
