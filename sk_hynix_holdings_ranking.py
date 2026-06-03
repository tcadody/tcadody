"""
SK하이닉스 임원·주요주주 주식 보유현황 랭킹 조회
DART OpenAPI (https://opendart.fss.or.kr) 사용

필요 환경변수:
  DART_API_KEY - DART OpenAPI 인증키 (https://opendart.fss.or.kr 에서 발급)

사용법 (Windows cmd):
  set DART_API_KEY=발급받은키
  python sk_hynix_holdings_ranking.py
"""

import os
import sys
import requests
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DART_BASE_URL = "https://opendart.fss.or.kr/api"
SK_HYNIX_CORP_CODE = "00164779"

# ──────────────────────────────────────────────
# DART API 호출
# ──────────────────────────────────────────────

def _get(endpoint: str, params: dict) -> dict:
    url = f"{DART_BASE_URL}/{endpoint}.json"
    resp = requests.get(url, params=params, timeout=15)
    resp.raise_for_status()
    return resp.json()


def fetch_elestock(api_key: str, bgn_de: str, end_de: str) -> list[dict]:
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

def _to_int(val) -> int:
    if not val or str(val).strip() in ("-", ""):
        return 0
    try:
        return int(str(val).replace(",", "").strip())
    except ValueError:
        return 0


def _to_float(val) -> float:
    if not val or str(val).strip() in ("-", ""):
        return 0.0
    try:
        return float(str(val).replace(",", "").strip())
    except ValueError:
        return 0.0


def build_ranking(records: list[dict], as_of_date: str = None) -> list[dict]:
    """
    as_of_date(YYYY-MM-DD) 기준 각 인물의 보유량을 추정.

    케이스 A: 기준일 이전 공시가 있으면 → 그 중 최신 공시의 sp_stock_lmp_cnt 사용
    케이스 B: 기준일 이전 공시가 없으면 → 기준일 이후 첫 공시에서
              sp_stock_lmp_cnt - sp_stock_lmp_irds_cnt 로 역산
    """
    # 인물별 전체 공시 목록 수집
    person_records: dict[str, list[dict]] = {}
    for r in records:
        nm = r.get("repror", "").strip()
        if not nm:
            continue
        person_records.setdefault(nm, []).append(r)

    result = []
    for nm, recs in person_records.items():
        recs_sorted = sorted(recs, key=lambda x: x.get("rcept_dt", ""))

        if as_of_date:
            before = [r for r in recs_sorted if r.get("rcept_dt", "") <= as_of_date]
            after  = [r for r in recs_sorted if r.get("rcept_dt", "") >  as_of_date]
        else:
            before = recs_sorted
            after  = []

        if before:
            # 케이스 A: 기준일 이전 최신 공시 그대로 사용
            ref = before[-1]
            holdings = _to_int(ref.get("sp_stock_lmp_cnt", 0))
            note = ref.get("rcept_dt", "-")
        elif after:
            # 케이스 B: 기준일 이후 첫 공시에서 역산
            ref = after[0]
            holdings = _to_int(ref.get("sp_stock_lmp_cnt", 0)) \
                     - _to_int(ref.get("sp_stock_lmp_irds_cnt", 0))
            note = f"{ref.get('rcept_dt','-')} 역산"
        else:
            continue

        result.append({
            "성명":       nm,
            "직위":       ref.get("isu_exctv_ofcps", "-").strip(),
            "등기여부":   ref.get("isu_exctv_rgist_at", "-").strip(),
            "주요주주":   ref.get("isu_main_shrholdr", "-").strip(),
            "보유주식수": max(holdings, 0),
            "증감":       _to_int(ref.get("sp_stock_lmp_irds_cnt", 0)),
            "보유비율":   _to_float(ref.get("sp_stock_lmp_rate", 0)),
            "최근보고일": note,
        })

    ranked = sorted(result, key=lambda x: x["보유주식수"], reverse=True)
    for i, row in enumerate(ranked, 1):
        row["순위"] = i
    return ranked


# ──────────────────────────────────────────────
# 콘솔 출력
# ──────────────────────────────────────────────

def print_ranking(ranking: list[dict], period: str):
    sep = "─" * 90
    print(f"\n{'═'*90}")
    print(f"  SK하이닉스 임원·주요주주 주식 보유현황 랭킹")
    print(f"  기준: {period}")
    print(f"{'═'*90}")
    print(f"{'순위':>4}  {'성명':<10} {'직위':<12} {'등기여부':<12} {'보유주식수(주)':>14} {'증감(주)':>12} {'보유비율':>8}  {'최근보고일'}")
    print(sep)

    for row in ranking:
        sign = "+" if row["증감"] >= 0 else ""
        print(
            f"{row['순위']:>4}  {row['성명']:<10} {row['직위']:<12} {row['등기여부']:<12} "
            f"{row['보유주식수']:>14,}  {sign}{row['증감']:>10,}  {row['보유비율']:>7.2f}%  {row['최근보고일']}"
        )

    print(sep)
    print(f"  총 {len(ranking)}명\n")


def search_person(all_rows: list[dict], name: str, bgn: str = "2023-01-01", end: str = "2024-12-31"):
    """특정 인물의 공시 이력을 기간 내에서 전부 출력."""
    hits = [
        r for r in all_rows
        if name in r.get("repror", "") and bgn <= r.get("rcept_dt", "") <= end
    ]
    sep = "─" * 70
    print(f"\n{'═'*70}")
    print(f"  '{name}' 공시 이력  ({bgn} ~ {end})")
    print(f"{'═'*70}")
    if not hits:
        print(f"  해당 기간에 '{name}' 공시 없음")
        # 전체 기간에서도 확인
        all_hits = [r for r in all_rows if name in r.get("repror", "")]
        if all_hits:
            dates = sorted(set(r.get("rcept_dt", "") for r in all_hits))
            print(f"  ※ 전체 이력에는 존재: {dates}")
        else:
            print(f"  ※ 전체 이력에도 없음 (이름 철자 확인 필요)")
    else:
        hits_sorted = sorted(hits, key=lambda x: x.get("rcept_dt", ""))
        print(f"  {'접수일':<12} {'직위':<12} {'등기여부':<12} {'보유주식수':>12} {'증감':>10}")
        print(sep)
        for r in hits_sorted:
            print(
                f"  {r.get('rcept_dt',''):<12} "
                f"{r.get('isu_exctv_ofcps',''):<12} "
                f"{r.get('isu_exctv_rgist_at',''):<12} "
                f"{_to_int(r.get('sp_stock_lmp_cnt',0)):>12,} "
                f"{_to_int(r.get('sp_stock_lmp_irds_cnt',0)):>+10,}"
            )
    print(sep + "\n")




HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
RANK1_FILL  = PatternFill("solid", fgColor="FFD700")
RANK2_FILL  = PatternFill("solid", fgColor="C0C0C0")
RANK3_FILL  = PatternFill("solid", fgColor="CD7F32")
ALT_FILL    = PatternFill("solid", fgColor="EBF3FB")
THIN   = Side(style="thin", color="BFBFBF")
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


def save_to_excel(ranking: list[dict], period: str, path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "SK하이닉스 주식보유랭킹"
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    ws.freeze_panes = "A4"

    # 제목
    ws.merge_cells("A1:I1")
    tc = ws["A1"]
    tc.value = f"SK하이닉스 임원·주요주주 주식 보유현황 랭킹  ({period})"
    tc.font = Font(bold=True, size=14, color="FFFFFF")
    tc.fill = PatternFill("solid", fgColor="1F4E79")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    tc.border = BORDER

    # 부제목
    ws.merge_cells("A2:I2")
    sc = ws["A2"]
    sc.value = f"조회일: {datetime.today().strftime('%Y-%m-%d')}  |  출처: DART 전자공시시스템 (opendart.fss.or.kr)"
    sc.font = Font(size=10, color="595959")
    sc.alignment = Alignment(horizontal="center", vertical="center")
    sc.border = BORDER

    # 헤더
    headers    = ["순위", "성명", "직위", "등기여부", "주요주주", "보유주식수(주)", "증감(주)", "보유비율(%)", "최근보고일"]
    col_widths = [7, 12, 14, 12, 10, 18, 14, 12, 14]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        _cell(ws, 3, col, h, bold=True, fill=HEADER_FILL, align="center")
        ws.column_dimensions[get_column_letter(col)].width = w

    # 데이터
    MEDAL = {1: RANK1_FILL, 2: RANK2_FILL, 3: RANK3_FILL}
    for row_idx, row in enumerate(ranking, 4):
        fill = MEDAL.get(row["순위"], ALT_FILL if row_idx % 2 == 0 else None)
        _cell(ws, row_idx, 1, row["순위"],       align="center", fill=fill)
        _cell(ws, row_idx, 2, row["성명"],        fill=fill)
        _cell(ws, row_idx, 3, row["직위"],        fill=fill)
        _cell(ws, row_idx, 4, row["등기여부"],    fill=fill)
        _cell(ws, row_idx, 5, row["주요주주"],    fill=fill)
        _cell(ws, row_idx, 6, row["보유주식수"],  bold=True, align="right", fill=fill, num_fmt="#,##0")
        _cell(ws, row_idx, 7, row["증감"],        align="right", fill=fill, num_fmt="+#,##0;-#,##0;0")
        _cell(ws, row_idx, 8, row["보유비율"],    align="right", fill=fill, num_fmt="0.0000%")
        _cell(ws, row_idx, 9, row["최근보고일"],  align="center", fill=fill)
        ws.row_dimensions[row_idx].height = 20

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
        print("  설정: set DART_API_KEY=발급받은키  (Windows cmd)")
        sys.exit(1)

    today  = datetime.today()
    end_de = today.strftime("%Y%m%d")

    # API가 날짜 파라미터를 무시하므로 전체 이력을 받아 build_ranking 에서 필터링
    AS_OF = "2024-12-31"  # 이 날짜 이전 마지막 공시 기준으로 스냅샷 생성
    period = f"2024-12-31 기준 (해당 시점까지의 최신 공시 보유수량)"

    print(f"  전체 이력 조회 중...", end=" ", flush=True)
    try:
        all_rows = fetch_elestock(api_key, "20000101", end_de)
        print(f"{len(all_rows)}건 수신")
    except Exception as e:
        print(f"실패 ({e})")
        sys.exit(1)

    records = all_rows

    # ── 전체 인물 목록 출력 (데이터 확인용) ──────────────────
    print("\n" + "═"*70)
    print("  데이터에 포함된 전체 인물 목록 (rcept_dt 범위)")
    print("═"*70)
    from collections import defaultdict
    person_dates: dict[str, list[str]] = defaultdict(list)
    for r in all_rows:
        nm = r.get("repror", "").strip()
        dt = r.get("rcept_dt", "")
        if nm:
            person_dates[nm].append(dt)
    for nm in sorted(person_dates):
        dates = sorted(person_dates[nm])
        print(f"  {nm:<12} {dates[0]} ~ {dates[-1]}  ({len(dates)}건)")
    print("═"*70 + "\n")

    # ── 특정 인물 검색 ──────────────────────────────────────
    search_person(all_rows, "임창문", bgn="2022-01-01", end="2024-12-31")

    if not records:
        print("조회 가능한 데이터가 없습니다.")
        sys.exit(1)

    ranking = build_ranking(records, as_of_date=AS_OF)

    if not ranking:
        print("랭킹 데이터를 생성할 수 없습니다.")
        sys.exit(1)

    print_ranking(ranking, period)

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "sk_hynix_holdings_2024.xlsx",
    )
    save_to_excel(ranking, period, out_path)


if __name__ == "__main__":
    main()
