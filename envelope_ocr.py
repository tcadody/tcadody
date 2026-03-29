#!/usr/bin/env python3
"""
봉투 OCR → 엑셀 자동 기록 시스템
수기로 이름과 금액이 적힌 봉투 이미지를 인식하여 엑셀 파일에 자동 기록합니다.

사용법:
  python envelope_ocr.py                      # 카메라 캡처 모드
  python envelope_ocr.py image.jpg            # 이미지 파일 모드
  python envelope_ocr.py *.jpg                # 여러 이미지 일괄 처리
  python envelope_ocr.py --output 결과.xlsx   # 출력 파일 지정
"""

import os
import re
import sys
import base64
import argparse
from datetime import datetime
from pathlib import Path

import anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── 설정 ──────────────────────────────────────────────────────────────────────

DEFAULT_OUTPUT = "봉투목록.xlsx"
CLAUDE_MODEL   = "claude-haiku-4-5-20251001"   # 빠르고 저렴한 Vision 모델

OCR_PROMPT = """이 봉투 이미지를 분석하여 수기로 적힌 이름과 금액을 추출해 주세요.

응답은 반드시 아래 형식으로만 답하세요 (다른 설명 없이):
이름: <추출한 이름>
금액: <숫자만, 예: 50000>

- 이름을 찾을 수 없으면: 이름: 알 수 없음
- 금액을 찾을 수 없으면: 금액: 0
- 금액은 쉼표, 원(₩), 만 등의 단위를 제거하고 숫자만 입력
- 예) "5만원" → 50000, "100,000원" → 100000, "10만" → 100000"""


# ─── 엑셀 헬퍼 ────────────────────────────────────────────────────────────────

def _thin_border():
    side = Side(style="thin")
    return Border(left=side, right=side, top=side, bottom=side)


def setup_excel(path: str) -> tuple[openpyxl.Workbook, object]:
    """엑셀 파일을 열거나 새로 만들고 시트·헤더를 준비합니다."""
    if Path(path).exists():
        wb = openpyxl.load_workbook(path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "봉투목록"

        # 헤더 행
        headers = ["번호", "날짜", "시간", "이름", "금액(원)", "이미지파일", "비고"]
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)

        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill   = header_fill
            cell.font   = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = _thin_border()

        # 열 너비
        widths = [6, 12, 10, 16, 14, 30, 20]
        for col, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(col)].width = width

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

    return wb, ws


def append_row(ws, name: str, amount: int, image_file: str = "", note: str = ""):
    """데이터 행을 추가합니다."""
    now    = datetime.now()
    # 다음 번호 계산 (헤더 제외)
    row_no = ws.max_row  # 헤더 포함이므로 max_row = 다음 번호

    data_font = Font(name="맑은 고딕", size=10)
    even_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    row_fill  = even_fill if row_no % 2 == 0 else None

    values = [
        row_no,
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S"),
        name,
        amount,
        image_file,
        note,
    ]

    new_row = ws.max_row + 1
    for col, val in enumerate(values, start=1):
        cell = ws.cell(row=new_row, column=col, value=val)
        cell.font   = data_font
        cell.border = _thin_border()
        cell.alignment = Alignment(vertical="center",
                                   horizontal="right" if col == 5 else "center")
        if row_fill:
            cell.fill = row_fill

    # 금액 열: 숫자 서식
    ws.cell(row=new_row, column=5).number_format = "#,##0"


def update_summary(ws):
    """합계 행을 마지막에 갱신합니다 (기존 합계 행 제거 후 재작성)."""
    # 기존 합계 행 찾아서 제거
    for row in range(ws.max_row, 1, -1):
        if ws.cell(row=row, column=4).value == "합 계":
            ws.delete_rows(row)
            break

    sum_row = ws.max_row + 1
    total_rows = sum_row - 2  # 헤더 + 마지막 합계 행 제외

    summary_font = Font(name="맑은 고딕", bold=True, size=10, color="FFFFFF")
    summary_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    labels = ["", "", "", "합 계", f"=SUM(E2:E{sum_row-1})", f"총 {total_rows}건", ""]
    for col, val in enumerate(labels, start=1):
        cell = ws.cell(row=sum_row, column=col, value=val)
        cell.font   = summary_font
        cell.fill   = summary_fill
        cell.border = _thin_border()
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if col == 5:
            cell.number_format = "#,##0"


# ─── OCR ──────────────────────────────────────────────────────────────────────

def image_to_base64(image_path: str) -> tuple[str, str]:
    """이미지 파일을 base64로 인코딩합니다."""
    suffix = Path(image_path).suffix.lower()
    media_types = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png",  ".gif": "image/gif",
        ".webp": "image/webp",
    }
    media_type = media_types.get(suffix, "image/jpeg")
    with open(image_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, media_type


def parse_ocr_response(text: str) -> tuple[str, int]:
    """Claude 응답에서 이름과 금액을 파싱합니다."""
    name   = "알 수 없음"
    amount = 0

    for line in text.strip().splitlines():
        line = line.strip()
        if line.startswith("이름:"):
            name = line.split(":", 1)[1].strip() or "알 수 없음"
        elif line.startswith("금액:"):
            raw = line.split(":", 1)[1].strip()
            digits = re.sub(r"[^\d]", "", raw)
            amount = int(digits) if digits else 0

    return name, amount


def ocr_envelope(image_path: str, client: anthropic.Anthropic) -> tuple[str, int, str]:
    """
    이미지에서 이름·금액을 추출합니다.
    반환: (이름, 금액, 원본_응답)
    """
    img_data, media_type = image_to_base64(image_path)

    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=256,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": img_data,
                        },
                    },
                    {"type": "text", "text": OCR_PROMPT},
                ],
            }
        ],
    )

    raw = message.content[0].text
    name, amount = parse_ocr_response(raw)
    return name, amount, raw


# ─── 카메라 캡처 ──────────────────────────────────────────────────────────────

def capture_from_camera(save_dir: str = ".") -> str | None:
    """
    웹캠으로 봉투를 촬영하고 임시 파일로 저장합니다.
    Space 또는 Enter: 촬영  |  q 또는 ESC: 취소
    """
    try:
        import cv2
    except ImportError:
        print("카메라 기능을 사용하려면 opencv-python을 설치하세요:")
        print("  pip install opencv-python")
        return None

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("카메라를 열 수 없습니다.")
        return None

    print("\n[카메라 모드]")
    print("  Space / Enter : 촬영")
    print("  q / ESC       : 취소\n")

    saved_path = None
    while True:
        ret, frame = cap.read()
        if not ret:
            break

        # 안내 텍스트 오버레이
        cv2.putText(frame, "Space/Enter: capture  |  q/ESC: quit",
                    (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
        cv2.imshow("봉투 촬영 - Envelope OCR", frame)

        key = cv2.waitKey(1) & 0xFF
        if key in (ord(" "), 13):   # Space 또는 Enter
            ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = str(Path(save_dir) / f"envelope_{ts}.jpg")
            cv2.imwrite(path, frame)
            print(f"사진 저장됨: {path}")
            saved_path = path
            break
        elif key in (ord("q"), 27):  # q 또는 ESC
            print("촬영 취소됨.")
            break

    cap.release()
    cv2.destroyAllWindows()
    return saved_path


# ─── 메인 ─────────────────────────────────────────────────────────────────────

def confirm(name: str, amount: int) -> tuple[str, int]:
    """인식 결과를 사용자에게 보여주고 수정 여부를 묻습니다."""
    print(f"\n  인식 결과 → 이름: {name}  /  금액: {amount:,}원")
    answer = input("  저장하시겠습니까? [Y/n/수정(e)] ").strip().lower()

    if answer == "n":
        return None, None  # 건너뛰기

    if answer == "e":
        new_name = input(f"  이름 수정 [{name}]: ").strip() or name
        raw_amt  = input(f"  금액 수정 [{amount}]: ").strip()
        digits   = re.sub(r"[^\d]", "", raw_amt)
        new_amt  = int(digits) if digits else amount
        return new_name, new_amt

    return name, amount


def process_image(image_path: str, client: anthropic.Anthropic,
                  wb: openpyxl.Workbook, ws, output: str,
                  auto: bool = False) -> bool:
    """단일 이미지를 처리합니다. 성공하면 True 반환."""
    print(f"\n처리 중: {image_path}")

    try:
        name, amount, raw = ocr_envelope(image_path, client)
    except Exception as e:
        print(f"  OCR 오류: {e}")
        return False

    if auto:
        print(f"  → 이름: {name}  /  금액: {amount:,}원")
        final_name, final_amount = name, amount
    else:
        final_name, final_amount = confirm(name, amount)
        if final_name is None:
            print("  건너뜀.")
            return False

    append_row(ws, final_name, final_amount, image_file=Path(image_path).name)
    update_summary(ws)
    wb.save(output)
    print(f"  저장 완료 → {output}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="봉투 이미지 OCR → 엑셀 자동 기록",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("images", nargs="*",
                        help="처리할 이미지 파일 경로 (생략 시 카메라 모드)")
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help=f"엑셀 출력 파일 (기본값: {DEFAULT_OUTPUT})")
    parser.add_argument("--auto", action="store_true",
                        help="확인 없이 자동 저장")
    parser.add_argument("--camera-loop", action="store_true",
                        help="카메라 반복 모드 (촬영 후 계속 찍기)")
    args = parser.parse_args()

    # API 키 확인
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("오류: ANTHROPIC_API_KEY 환경 변수가 설정되지 않았습니다.")
        print("  export ANTHROPIC_API_KEY='sk-ant-...'")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)
    wb, ws = setup_excel(args.output)

    # ── 이미지 파일 모드 ──
    if args.images:
        count = 0
        for pattern in args.images:
            paths = sorted(Path(".").glob(pattern)) if "*" in pattern else [Path(pattern)]
            for p in paths:
                if not p.exists():
                    print(f"파일 없음: {p}")
                    continue
                if process_image(str(p), client, wb, ws, args.output, auto=args.auto):
                    count += 1
        print(f"\n완료: 총 {count}건 저장 → {args.output}")
        return

    # ── 카메라 모드 ──
    print("=== 봉투 OCR 카메라 모드 ===")
    print(f"저장 파일: {args.output}")
    count = 0

    while True:
        image_path = capture_from_camera()
        if image_path is None:
            break

        if process_image(image_path, client, wb, ws, args.output, auto=args.auto):
            count += 1

        if not args.camera_loop:
            again = input("\n계속 촬영하시겠습니까? [Y/n] ").strip().lower()
            if again == "n":
                break

    print(f"\n완료: 총 {count}건 저장 → {args.output}")


if __name__ == "__main__":
    main()
