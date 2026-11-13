#!/usr/bin/env python3
"""
위임장 선별 판독기

PDF 묶음(위임장, 의결권행사서, 원본증명서, 감사추적인증서 등이 섞인 파일)에서
'제목에 위임장이 들어간 페이지'만 골라 항목을 판독하고 엑셀 목록으로 만듭니다.

사용법:
    python weimjang_reader.py <PDF 파일 또는 폴더> [-o 결과.xlsx]
                              [--keyword 위임장] [--no-ocr]
                              [--manual manual_overrides.json]

동작 순서:
  1. 페이지마다 텍스트를 추출합니다. 텍스트가 없는 스캔·사진 페이지는 OCR(Tesseract kor+eng)로 읽습니다.
  2. 페이지 '제목'(머리말을 제외한 첫 줄)에 키워드(기본: 위임장)가 들어간 페이지만 선별합니다.
     → 본문에 '위임장'이 언급되기만 한 원본증명서·감사추적인증서·별지 등은 제외됩니다.
  3. 선별된 위임장에서 위임인 성명·생년월일·주소·연락처·호수·작성일·수임인 등을 판독합니다.
  4. 같은 일련번호의 감사추적인증서가 있으면 전자서명 완료일시·발송자를 함께 기록합니다(대조용).
  5. 필수 항목 누락, 수임인 불일치, 호수 중복, 수기(OCR) 판독 페이지를 '검토필요'로 표시합니다.
  6. 첫 시트에 이름·호수·전화번호·수임인만 호수 순으로 정리합니다 (--contacts-only 로 이 시트만 출력).

수기 작성 위임장은 OCR 정확도가 낮습니다. 사람이 확인한 값을 --manual JSON으로 넣으면
해당 페이지 값이 덮어써지고 판독방식이 '육안판독(수동입력)'으로 기록됩니다.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------
HEADER_LINE = re.compile(r"^\(일련번호\s*\d+\)|^\[별지\]$|^\[\d{4}-\d{2}-\d{2}\]")
SERIAL = re.compile(r"\(일련번호\s*(\d+)\)\s*\[(\d{4}-\d{2}-\d{2})\]")
GUIDE = re.compile(r"서명|날인|법인|인장|해당란|기재")
OCR_TITLE_CROP = 0.12  # 스캔 페이지에서 제목을 찾을 상단 영역 비율

COLUMNS = [
    ("No", 5), ("PDF", 22), ("페이지", 7), ("문서제목", 30), ("판독방식", 16),
    ("위임인구분", 12), ("성명", 12), ("생년월일", 12), ("주소", 40), ("연락처", 15),
    ("소유(점유)호수", 16), ("작성일", 12), ("서명/날인", 10), ("수임인", 12),
    ("수임인 정보", 40), ("일련번호", 10), ("발급일", 12),
    ("전자서명 완료일시", 19), ("발송자", 24), ("검토필요", 34), ("페이지 이미지", 20),
]


# ---------------------------------------------------------------------------
# 텍스트 추출
# ---------------------------------------------------------------------------
def page_texts(pdf_path: Path) -> list[str]:
    import pdfplumber

    with pdfplumber.open(str(pdf_path)) as pdf:
        return [(p.extract_text() or "") for p in pdf.pages]


def render_page(pdf_path: Path, page_no: int, dpi: int = 300):
    from pdf2image import convert_from_path

    return convert_from_path(str(pdf_path), dpi=dpi, first_page=page_no, last_page=page_no)[0]


def ocr(image) -> str:
    import pytesseract

    return pytesseract.image_to_string(image, lang="kor+eng")


def ocr_available() -> bool:
    try:
        import pytesseract

        pytesseract.get_tesseract_version()
        return "kor" in pytesseract.get_languages(config="")
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 제목 판별
# ---------------------------------------------------------------------------
def title_of(text: str) -> str:
    for line in text.splitlines():
        line = line.strip()
        if not line or HEADER_LINE.search(line):
            continue
        return line
    return ""


# ---------------------------------------------------------------------------
# 항목 판독
# ---------------------------------------------------------------------------
def squash(s: str) -> str:
    """'신 영 희' → '신영희', '1 9 6 7' → '1967' (글자 사이 공백 제거)."""
    return re.sub(r"\s+", "", s or "")


def field(text: str, label: str, stop: str = r"$") -> str:
    """'라벨 : 값' 형태에서 값만 꺼냅니다. 라벨 글자 사이 공백을 허용합니다."""
    lab = r"\s*".join(map(re.escape, label))
    m = re.search(lab + r"\s*[:：]?\s*(.+?)\s*" + stop, text, flags=re.MULTILINE)
    return m.group(1).strip() if m else ""


def norm_date(s: str) -> str:
    digits = re.findall(r"\d+", squash(s))
    joined = "".join(digits)
    if len(digits) >= 3:
        y, m, d = digits[0], digits[1], digits[2]
    elif len(joined) == 8:
        y, m, d = joined[:4], joined[4:6], joined[6:]
    else:
        return squash(s)
    try:
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    except ValueError:
        return squash(s)


def parse_weimjang(text: str) -> dict:
    rec: dict[str, str] = {}
    rec["성명"] = squash(field(text, "성명", r"(?:\(서명|\(인|$)"))
    if GUIDE.search(rec["성명"]) or len(rec["성명"]) > 20:  # 양식 안내 문구는 이름이 아님
        rec["성명"] = ""
    rec["생년월일"] = norm_date(field(text, "생년월일", r"(?:\(|$)"))
    rec["주소"] = re.sub(r"\s+", " ", field(text, "주소"))
    rec["연락처"] = squash(field(text, "연락처"))
    rec["소유(점유)호수"] = ", ".join(
        f"{h}호" for h in re.findall(r"(\d{3,4})\s*호", field(text, "소유(점유)호수"))
    )
    rec["작성일"] = norm_date(field(text, "작성일"))
    if "(서명/인)" in text or "(서명" in text:
        rec["서명/날인"] = "서명란 있음"

    if re.search(r"위임인\s*\(\s*소유자\s*\)", text):
        rec["위임인구분"] = "구분소유자"
    elif re.search(r"위임인\s*\(\s*점유자\s*\)", text):
        rec["위임인구분"] = "점유자"

    # 수임인: '수임인(대리인)' 다음 줄의 '이름 / 생년월일 / 주소'
    m = re.search(r"수임인\s*\(대리인\)\s*\n\s*([^\n]+)", text)
    if m and "/" in m.group(1):
        info = m.group(1).strip()
        rec["수임인 정보"] = info
        rec["수임인"] = squash(info.split("/")[0])
    return rec


def parse_audit(text: str) -> dict:
    """같은 일련번호의 감사추적인증서에서 전자서명 완료일시·발송자를 찾습니다."""
    out = {}
    m = re.search(r"\((\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\)\s*수신자\s*(\S+?)님이 문서를 서명 완료", text)
    if m:
        out["전자서명 완료일시"] = m.group(1)
        out["서명자(감사추적)"] = m.group(2)
    m = re.search(r"^(.+관리인.+)$", text, flags=re.MULTILINE)
    if m:
        out["발송자"] = m.group(1).strip()
    return out


# ---------------------------------------------------------------------------
# 메인 처리
# ---------------------------------------------------------------------------
def process_pdf(pdf_path: Path, keyword: str, use_ocr: bool, img_dir: Path | None):
    texts = page_texts(pdf_path)
    pages, records = [], []
    audits: dict[str, dict] = {}

    for idx, text in enumerate(texts, start=1):
        method = "텍스트"
        full_text = text
        if not text.strip():
            if use_ocr:
                img = render_page(pdf_path, idx)
                w, h = img.size
                title = title_of(ocr(img.crop((0, 0, w, int(h * OCR_TITLE_CROP)))))
                method = "OCR"
            else:
                img, title = None, ""
                method = "이미지(OCR 꺼짐)"
        else:
            img = None
            title = title_of(text)

        serial = SERIAL.search(text)
        if "감사" in title or "수신자" in text and "서명 완료" in text:
            if serial:
                audits[serial.group(1)] = parse_audit(text)

        selected = keyword in title
        pages.append({"PDF": pdf_path.name, "페이지": idx, "제목": title or "(제목 미확인)",
                      "판독방식": method, "선별": "예" if selected else "아니오"})
        if not selected:
            continue

        if method == "OCR":
            full_text = ocr(img)
        rec = parse_weimjang(full_text)
        rec.update({"PDF": pdf_path.name, "페이지": idx, "문서제목": title, "판독방식": method})
        if serial:
            rec["일련번호"], rec["발급일"] = serial.group(1), serial.group(2)
        if method == "OCR":
            if img_dir:
                img_dir.mkdir(parents=True, exist_ok=True)
                p = img_dir / f"{pdf_path.stem}_p{idx:03d}.png"
                img.save(p)
                rec["페이지 이미지"] = str(p)
            rec["_ocr"] = True
        records.append(rec)

    # 감사추적 정보 연결 (일련번호 기준)
    for rec in records:
        a = audits.get(rec.get("일련번호", ""))
        if a:
            rec["전자서명 완료일시"] = a.get("전자서명 완료일시", "")
            rec["발송자"] = a.get("발송자", "")
            if a.get("서명자(감사추적)") and not rec.get("성명"):
                rec["성명"] = a["서명자(감사추적)"]
                rec.setdefault("_issues", []).append("이름은 감사추적인증서 서명자에서 가져옴")
            elif a.get("서명자(감사추적)") and rec.get("성명") and a["서명자(감사추적)"] != rec["성명"]:
                rec.setdefault("_issues", []).append(f"감사추적 서명자({a['서명자(감사추적)']})와 성명 불일치")
    return pages, records


def apply_manual(records: list[dict], manual: dict) -> None:
    """manual: {"<PDF파일명>:<페이지>" 또는 "<페이지>": {열이름: 값, ...}}"""
    for rec in records:
        key_full, key_page = f"{rec['PDF']}:{rec['페이지']}", str(rec["페이지"])
        override = manual.get(key_full) or manual.get(key_page)
        if override:
            rec.update({k: v for k, v in override.items() if not k.startswith("_")})
            rec["판독방식"] = "육안판독(수동입력)"
            rec["_ocr"] = False
            if override.get("_비고"):
                rec.setdefault("_issues", []).append(override["_비고"])


def validate(records: list[dict]) -> None:
    agents = Counter(r.get("수임인") for r in records if r.get("수임인"))
    main_agent = agents.most_common(1)[0][0] if agents else ""
    units = Counter()
    for r in records:
        for u in filter(None, (s.strip() for s in r.get("소유(점유)호수", "").split(","))):
            units[u] += 1

    for r in records:
        issues = r.setdefault("_issues", [])
        if r.get("_ocr"):
            issues.insert(0, "수기/스캔(OCR) – 원본 대조 필요")
        missing = [k for k in ("성명", "소유(점유)호수", "작성일", "수임인") if not r.get(k)]
        if missing:
            issues.append("누락: " + ", ".join(missing))
        if main_agent and r.get("수임인") and r["수임인"] != main_agent:
            issues.append(f"수임인 상이({r['수임인']})")
        dup = [u for u in r.get("소유(점유)호수", "").split(", ") if u and units[u] > 1]
        if dup:
            issues.append("호수 중복: " + ", ".join(dup))
        r["검토필요"] = " / ".join(dict.fromkeys(issues)) if issues else ""


# ---------------------------------------------------------------------------
# 엑셀 출력
# ---------------------------------------------------------------------------
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")


def style_header(ws, widths):
    for i, (name, width) in enumerate(widths, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(widths))}1"


def format_phone(s: str) -> str:
    d = re.sub(r"\D", "", s or "")
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    if len(d) == 10:
        return f"{d[:3]}-{d[3:6]}-{d[6:]}"
    return s or ""


CONTACT_COLS = [("No", 5), ("이름", 12), ("호수", 18), ("전화번호", 16), ("수임인", 12), ("비고", 30)]


def write_contacts(ws, records) -> None:
    """이름·호수·전화번호 요약 시트 (호수 오름차순)."""
    style_header(ws, CONTACT_COLS)

    def first_unit(r):
        m = re.search(r"\d+", r.get("소유(점유)호수", ""))
        return int(m.group()) if m else 99999

    for n, r in enumerate(sorted(records, key=first_unit), start=1):
        note = "원본 대조 필요" if r.get("_ocr") or r.get("판독방식", "").startswith("육안") else ""
        vals = [n, r.get("성명", ""), r.get("소유(점유)호수", ""), format_phone(r.get("연락처", "")),
                r.get("수임인", ""), note]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=n + 1, column=c, value=v)
            cell.alignment = Alignment(vertical="top")
            if c == 4:
                cell.number_format = "@"  # 앞자리 0 보존


def write_excel(records, pages, out: Path, contacts_only: bool = False) -> None:
    wb = Workbook()
    write_contacts(wb.active, records)
    wb.active.title = "이름·호수·전화번호"
    if contacts_only:
        wb.save(out)
        return

    ws = wb.create_sheet("위임장 목록")
    style_header(ws, COLUMNS)
    for n, r in enumerate(records, start=1):
        row = n + 1
        for c, (name, _) in enumerate(COLUMNS, start=1):
            val = n if name == "No" else r.get(name, "")
            cell = ws.cell(row=row, column=c, value=val)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if name == "페이지 이미지" and val:
                cell.hyperlink = val
                cell.value = Path(val).name
                cell.font = Font(color="0563C1", underline="single")
        if r.get("검토필요"):
            for c in range(1, len(COLUMNS) + 1):
                ws.cell(row=row, column=c).fill = WARN_FILL

    ps = wb.create_sheet("페이지 분류")
    pcols = [("PDF", 22), ("페이지", 7), ("제목", 50), ("판독방식", 16), ("선별", 8)]
    style_header(ps, pcols)
    for i, p in enumerate(pages, start=2):
        for c, (name, _) in enumerate(pcols, start=1):
            ps.cell(row=i, column=c, value=p[name])

    sm = wb.create_sheet("요약")
    review = sum(1 for r in records if r.get("검토필요"))
    for row in (["항목", "값"], ["전체 페이지", len(pages)], ["위임장 선별", len(records)],
                ["검토 필요", review], ["이상 없음", len(records) - review]):
        sm.append(row)
    sm["A1"].font = sm["B1"].font = Font(bold=True)
    sm.column_dimensions["A"].width = 16
    wb.save(out)


def main() -> int:
    ap = argparse.ArgumentParser(description="제목에 '위임장'이 들어간 페이지만 선별해 판독 → 엑셀")
    ap.add_argument("target", help="PDF 파일 또는 폴더")
    ap.add_argument("-o", "--output", default="위임장_판독결과.xlsx")
    ap.add_argument("--keyword", default="위임장", help="제목 선별 키워드 (기본: 위임장)")
    ap.add_argument("--no-ocr", action="store_true", help="스캔 페이지 OCR 생략")
    ap.add_argument("--manual", help="수기 위임장 육안판독 값을 담은 JSON")
    ap.add_argument("--contacts-only", action="store_true",
                    help="이름·호수·전화번호 시트만 출력")
    args = ap.parse_args()

    target = Path(args.target)
    pdfs = sorted(target.rglob("*.pdf")) if target.is_dir() else [target]
    if not pdfs:
        print("PDF를 찾지 못했습니다.", file=sys.stderr)
        return 1

    use_ocr = not args.no_ocr and ocr_available()
    if not args.no_ocr and not use_ocr:
        print("※ Tesseract(kor)가 없어 스캔 페이지는 OCR 없이 건너뜁니다.", file=sys.stderr)

    out = Path(args.output)
    img_dir = out.with_name(out.stem + "_페이지이미지")
    all_pages, all_recs = [], []
    for pdf in pdfs:
        pages, recs = process_pdf(pdf, args.keyword, use_ocr, img_dir)
        all_pages += pages
        all_recs += recs

    if args.manual:
        apply_manual(all_recs, json.loads(Path(args.manual).read_text(encoding="utf-8")))
    validate(all_recs)
    write_excel(all_recs, all_pages, out, contacts_only=args.contacts_only)

    for r in all_recs:
        flag = f"  ⚠ {r['검토필요']}" if r.get("검토필요") else ""
        print(f"p{r['페이지']:>3} {r.get('성명', '?'):<6} {r.get('소유(점유)호수', ''):<16}{flag}")
    print(f"\n위임장 {len(all_recs)}건 / 전체 {len(all_pages)}페이지 → {out.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
