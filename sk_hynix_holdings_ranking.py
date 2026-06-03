"""
SK하이닉스 임원·주요주주 주식 보유현황 랭킹 조회
DART OpenAPI (https://opendart.fss.or.kr) 사용

필요 환경변수:
  DART_API_KEY - DART OpenAPI 인증키 (https://opendart.fss.or.kr 에서 발급)

사용법:
  export DART_API_KEY=your_key_here
  python sk_hynix_holdings_ranking.py
"""

import os
import sys
import requests
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────
DART_BASE_URL = "https://opendart.fss.or.kr/api"
SK_HYNIX_CORP_CODE = "00164779"  # SK하이닉스 고유번호

REPORT_CODES = {
    "11011": "사업보고서",
    "11012": "반기보고서",
    "11013": "1분기보고서",
    "11014": "3분기보고서",
}

# ──────────────────────────────────────────────
# DART API 호출
# ──────────────────────────────────────────────

def _get(endpoint: str, params: dict) -> dict:
    url = f"{DART_BASE_URL}/{endpoint}.json"
    resp = requests.get(url, params=params, timeout=15)
    resp.raise_for_status()
    return resp.json()


def fetch_elestock(api_key: str, bgn_de: str, end_de: str) -> list[dict]:
    """임원·주요주주 소유보고 (지분공시) 조회.
    bgn_de, end_de: YYYYMMDD 형식
    """
    data = _get(
        "elestock",
        {
            "crtfc_key": api_key,
            "corp_code": SK_HYNIX_CORP_CODE,
            "bgn_de": bgn_de,
            "end_de": end_de,
        },
    )
    status = data.get("status")
    if status == "013":
        return []
    if status != "000":
        raise RuntimeError(f"DART API 오류 [{status}]: {data.get('message')}")
    return data.get("list", [])


# ──────────────────────────────────────────────
# 데이터 처리
# ──────────────────────────────────────────────

def _to_int(val: str) -> int:
    if not val or val.strip() in ("-", ""):
        return 0
    return int(val.replace(",", "").strip())


def build_ranking(records: list[dict]) -> list[dict]:
    """
    elestock 보고 이력에서 인물+주식종류별 최신 보고를 골라 합산 후 랭킹 생성.
    동일 인물이 여러 번 신고한 경우 가장 최근 after_srn_cnt 를 사용.
    """
    # (성명, 주식종류) → 가장 최신 레코드 (change_on 기준)
    latest: dict[tuple, dict] = {}
    for r in records:
        nm = r.get("nm", "").strip()
        stock_knd = r.get("stock_knd", "").strip()
        change_on = r.get("change_on", "").strip()
        key = (nm, stock_knd)
        if key not in latest or change_on > latest[key].get("change_on", ""):
            latest[key] = r

    # 성명별 합산
    agg: dict[str, dict] = {}
    for (nm, stock_knd), r in latest.items():
        relate = r.get("relate", "").strip()
        after_cnt = _to_int(r.get("after_srn_cnt", "0"))
        before_cnt = _to_int(r.get("before_srn_cnt", "0"))
        incrs_cnt = _to_int(r.get("incrs_srn_cnt", "0"))
        dcrs_cnt = _to_int(r.get("dcrs_srn_cnt", "0"))
        change_on = r.get("change_on", "").strip()

        if nm not in agg:
            agg[nm] = {
                "성명": nm,
                "관계": relate,
                "보유주식수_합계": 0,
                "기초_합계": 0,
                "증가_합계": 0,
                "감소_합계": 0,
                "최근변동일": "",
                "주식종류_목록": [],
            }

        agg[nm]["보유주식수_합계"] += after_cnt
        agg[nm]["기초_합계"] += before_cnt
        agg[nm]["증가_합계"] += incrs_cnt
        agg[nm]["감소_합계"] += dcrs_cnt
        if change_on > agg[nm]["최근변동일"]:
            agg[nm]["최근변동일"] = change_on
        if stock_knd and stock_knd not in agg[nm]["주식종류_목록"]:
            agg[nm]["주식종류_목록"].append(stock_knd)

    ranked = sorted(agg.values(), key=lambda x: x["보유주식수_합계"], reverse=True)
    for i, row in enumerate(ranked, 1):
        row["순위"] = i
        row["주식종류"] = ", ".join(row["주식종류_목록"])
        row["순증감"] = row["증가_합계"] - row["감소_합계"]

    return ranked


# ──────────────────────────────────────────────
# 콘솔 출력
# ──────────────────────────────────────────────

def print_ranking(ranking: list[dict], year: str, reprt_label: str):
    sep = "─" * 80
    print(f"\n{'═'*80}")
    print(f"  SK하이닉스 임원·주요주주 주식 보유현황 랭킹")
    print(f"  기준: {year}년 {reprt_label}")
    print(f"{'═'*80}")
    header = f"{'순위':>4}  {'성명':<12} {'관계':<16} {'주식종류':<12} {'기말보유(주)':>15} {'순증감(주)':>14}"
    print(header)
    print(sep)

    for row in ranking:
        sign = "+" if row["순증감"] >= 0 else ""
        incr_str = f"{sign}{row['순증감']:,}"
        print(
            f"{row['순위']:>4}  {row['성명']:<12} {row['관계']:<16} "
            f"{row['주식종류']:<12} {row['보유주식수_합계']:>15,}  {incr_str:>14}"
        )

    print(sep)
    print(f"  총 {len(ranking)}명\n")


# ──────────────────────────────────────────────
# 엑셀 저장
# ──────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
RANK1_FILL = PatternFill("solid", fgColor="FFD700")
RANK2_FILL = PatternFill("solid", fgColor="C0C0C0")
RANK3_FILL = PatternFill("solid", fgColor="CD7F32")
ALT_FILL = PatternFill("solid", fgColor="EBF3FB")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _cell(ws, row, col, value, bold=False, fill=None, align="left", num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, color="FFFFFF" if fill == HEADER_FILL else "000000", size=11)
    if fill:
        c.fill = fill
    c.border = BORDER
    c.alignment = Alignment(horizontal=align, vertical="center")
    if num_fmt:
        c.number_format = num_fmt
    return c


def save_to_excel(ranking: list[dict], year: str, reprt_label: str, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "SK하이닉스 주식보유랭킹"
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = "A4"

    # 제목
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = f"SK하이닉스 임원·주요주주 주식 보유현황 랭킹  ({year}년 {reprt_label})"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = BORDER

    # 부제목
    ws.merge_cells("A2:H2")
    sub_cell = ws["A2"]
    sub_cell.value = f"조회일: {datetime.today().strftime('%Y-%m-%d')}  |  출처: DART 전자공시시스템"
    sub_cell.font = Font(size=10, color="595959")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    sub_cell.border = BORDER

    # 헤더
    headers = ["순위", "성명", "관계", "주식종류", "기초 보유(주)", "증가(주)", "감소(주)", "기말 보유(주)"]
    col_widths = [8, 14, 20, 16, 18, 16, 16, 18]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        _cell(ws, 3, col, h, bold=True, fill=HEADER_FILL, align="center")
        ws.column_dimensions[get_column_letter(col)].width = w

    # 데이터 행
    MEDAL = {1: RANK1_FILL, 2: RANK2_FILL, 3: RANK3_FILL}
    for row_idx, row in enumerate(ranking, 4):
        r = row_idx
        fill = MEDAL.get(row["순위"], ALT_FILL if row_idx % 2 == 0 else None)
        _cell(ws, r, 1, row["순위"], align="center", fill=fill)
        _cell(ws, r, 2, row["성명"], fill=fill)
        _cell(ws, r, 3, row["관계"], fill=fill)
        _cell(ws, r, 4, row["주식종류"], fill=fill)
        _cell(ws, r, 5, row["기초_합계"], align="right", fill=fill, num_fmt="#,##0")
        _cell(ws, r, 6, row["증가_합계"], align="right", fill=fill, num_fmt="#,##0")
        _cell(ws, r, 7, row["감소_합계"], align="right", fill=fill, num_fmt="#,##0")
        _cell(ws, r, 8, row["보유주식수_합계"], bold=True, align="right", fill=fill, num_fmt="#,##0")
        ws.row_dimensions[r].height = 20

    wb.save(path)
    print(f"엑셀 저장 완료: {path}")


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────

def main():
    api_key = os.getenv("DART_API_KEY", "").strip()
    if not api_key:
        print("오류: DART_API_KEY 환경변수가 설정되지 않았습니다.")
        print("  발급: https://opendart.fss.or.kr > 인증키 신청/관리")
        print("  설정: set DART_API_KEY=발급받은키  (Windows)")
        sys.exit(1)

    # 최근 3년치 소유보고 이력 조회
    today = datetime.today()
    end_de = today.strftime("%Y%m%d")
    bgn_de = today.replace(year=today.year - 3).strftime("%Y%m%d")
    label_year = str(today.year - 1)

    print(f"  SK하이닉스 임원·주요주주 소유보고 조회 중 ({bgn_de} ~ {end_de})...", end=" ", flush=True)
    try:
        records = fetch_elestock(api_key, bgn_de, end_de)
    except Exception as e:
        print(f"실패 ({e})")
        sys.exit(1)

    if not records:
        print("데이터 없음")
        print("조회된 소유보고 데이터가 없습니다. API 키 또는 조회 기간을 확인해주세요.")
        sys.exit(1)

    print(f"{len(records)}건 수신")

    ranking = build_ranking(records)
    print_ranking(ranking, label_year, "임원·주요주주 소유보고 (최근 3년)")

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        f"sk_hynix_holdings_{label_year}.xlsx",
    )
    save_to_excel(ranking, label_year, "임원·주요주주 소유보고", out_path)
    return out_path


if __name__ == "__main__":
    main()
