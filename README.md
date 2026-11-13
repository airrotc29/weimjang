# weimjang

위임장 관련 문서 도구 모음입니다.

| 파일 | 내용 |
|---|---|
| `weimjang_reader.py` | PDF 묶음에서 **제목에 '위임장'이 들어간 페이지만** 골라 판독하고 엑셀로 정리 |
| `power_of_attorney_en.md` | 영문 위임장(Power of Attorney) 템플릿 |
| `manual_overrides.example.json` | 수기 위임장 육안판독 값 입력 예시 |

## weimjang_reader.py

위임장, 의결권행사서, 원본증명서, 감사추적인증서, 별지 명단 등이 섞인 PDF에서 페이지 제목(머리말 제외 첫 줄)에 `위임장`이 있는 페이지만 선별합니다. 본문에만 '위임장'이 언급된 증명서 페이지는 제외됩니다.

### 설치

```bash
pip install -r requirements.txt
```

스캔·사진 페이지를 읽으려면 Tesseract OCR(한국어)도 필요합니다.

```bash
sudo apt-get install -y tesseract-ocr tesseract-ocr-kor poppler-utils
```

### 실행

```bash
python weimjang_reader.py 위임장묶음.pdf -o 위임장_판독결과.xlsx
```

- 폴더를 지정하면 하위 PDF를 모두 처리합니다.
- `--keyword 철회서`: 다른 제목 키워드로 선별
- `--contacts-only`: 이름·호수·전화번호·수임인 시트만 출력
- `--manual manual_overrides.json`: 수기 위임장 값을 사람이 확인한 값으로 덮어쓰기
- `--no-ocr`: 스캔 페이지 OCR 생략

### 엑셀 시트

1. **이름·호수·전화번호**: 이름, 호수, 전화번호, 수임인 (호수 순)
2. **위임장 목록**: 위임인구분, 생년월일, 주소, 작성일, 수임인 정보, 일련번호, 전자서명 완료일시(같은 일련번호의 감사추적인증서에서 연결), 검토필요
3. **페이지 분류**: 모든 페이지의 제목과 선별 여부
4. **요약**: 건수 집계

### 주의

- 수기 위임장은 OCR 정확도가 낮습니다. 엑셀의 페이지 이미지 링크로 원본을 확인하고 `--manual`로 값을 입력하세요.
- 판독 결과는 검토 자료입니다. 대리권 증명(집합건물법 제38조 제2항)의 유효성은 원본으로 확인해야 합니다.
- 결과 엑셀과 원본 PDF에는 개인정보가 들어 있으므로 저장소에 커밋하지 마세요 (`.gitignore`에 제외 설정됨).
