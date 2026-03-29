# 봉투 OCR → 엑셀 자동 기록

수기로 이름과 금액이 적힌 봉투 이미지를 카메라로 찍거나 이미지 파일로 입력하면,
Claude Vision AI가 텍스트를 인식하여 엑셀 파일에 자동으로 기록합니다.

## 설치

```bash
pip install -r requirements.txt

# 카메라 기능도 사용하려면
pip install opencv-python
```

## API 키 설정

```bash
# .env.example을 복사하여 실제 키 입력
cp .env.example .env

# 또는 직접 환경 변수 설정
export ANTHROPIC_API_KEY="sk-ant-..."
```

API 키는 [Anthropic Console](https://console.anthropic.com)에서 발급받으세요.

## 사용법

### 1. 카메라 촬영 모드 (기본)
```bash
python envelope_ocr.py
```
- Space 또는 Enter: 촬영
- q 또는 ESC: 종료
- 촬영 후 인식 결과 확인 → Y(저장) / n(건너뜀) / e(수정)

### 2. 이미지 파일 모드
```bash
python envelope_ocr.py 봉투사진.jpg
```

### 3. 여러 이미지 일괄 처리
```bash
python envelope_ocr.py 봉투1.jpg 봉투2.jpg 봉투3.jpg
python envelope_ocr.py "*.jpg"
```

### 4. 옵션
```bash
# 출력 파일 지정
python envelope_ocr.py -o 결혼식_봉투.xlsx 사진들/*.jpg

# 자동 저장 (확인 없이)
python envelope_ocr.py --auto *.jpg

# 카메라 반복 모드 (계속 촬영)
python envelope_ocr.py --camera-loop
```

## 엑셀 출력 형식

| 번호 | 날짜 | 시간 | 이름 | 금액(원) | 이미지파일 | 비고 |
|------|------|------|------|----------|-----------|------|
| 1 | 2024-01-01 | 14:30:00 | 홍길동 | 50,000 | envelope_001.jpg | |
| 2 | 2024-01-01 | 14:31:00 | 김철수 | 100,000 | envelope_002.jpg | |
| | | | **합 계** | **150,000** | 총 2건 | |

- 파일이 이미 존재하면 기존 데이터에 이어서 추가됩니다.
- 합계 행이 자동으로 갱신됩니다.

## 지원 이미지 형식

JPG, JPEG, PNG, GIF, WEBP
