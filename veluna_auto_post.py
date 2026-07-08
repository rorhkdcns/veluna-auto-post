# -*- coding: utf-8 -*-
"""
벨루나몰(velunamall.com) 상품 -> 블로그 자동 포스팅 스크립트

흐름:
1. Google Sheet에서 '포스팅완료'가 비어있는 상품 중 랜덤 1개 선택
2. 상품상세URL 스크래핑 -> 실제 상품 이미지(썸네일 + 상세이미지) 추출
3. 카테고리 -> 남성/여성/커플 분류 (category_map.py)
4. Gemini API로 포스팅 본문 생성
5. Blogger API로 포스팅 업로드 (썸네일 먼저, 그 다음 상세이미지, 그 다음 본문)
6. 포스팅 완료된 행에 '포스팅완료' 표시 (타임스탬프 + 포스팅 URL)

필요 환경변수 (GitHub Actions Secrets로 등록):
- GOOGLE_SHEETS_CREDENTIALS : 서비스 계정 JSON 전체 내용 (문자열)
- SPREADSHEET_ID            : 구글시트 ID
- GEMINI_API_KEY            : Gemini API 키
- BLOGGER_TOKEN_PICKLE_B64  : Blogger OAuth token.pickle을 base64 인코딩한 문자열
                              (car-auto-post의 TOKEN_PICKLE_BASE64와 동일한 방식)
- BLOGGER_BLOG_ID           : 포스팅할 Blogger 블로그 ID
"""

import os
import json
import random
import base64
import pickle
import io
import re
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from PIL import Image
import gspread
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from google import genai

from category_map import classify

# ---------- 카테고리별 고정 썸네일 ----------
THUMBNAIL_FILES = {
    "남성": "thumbnails/thumb_male.jpg",
    "여성": "thumbnails/thumb_female.jpg",
    "커플": "thumbnails/thumb_couple.jpg",
}


def get_category_thumbnail_url(category_label: str) -> str:
    repo = os.environ.get("GITHUB_REPOSITORY", "")
    branch = "main"
    path = THUMBNAIL_FILES.get(category_label, THUMBNAIL_FILES["커플"])
    return f"https://raw.githubusercontent.com/{repo}/{branch}/{path}"

# ---------- 설정 ----------
SHEET_TAB = "Sheet1"
POSTED_COL_NAME = "포스팅완료"
UA_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
}

# ---------- Google Sheets ----------
def get_sheet():
    creds_json = json.loads(os.environ["GOOGLE_SHEETS_CREDENTIALS"])
    creds = SACredentials.from_service_account_info(
        creds_json,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(os.environ["SPREADSHEET_ID"])
    return sh.sheet1  # 탭 이름과 무관하게 첫 번째 탭을 사용


EXPECTED_FIRST_HEADER = "인덱스(변경불가)"
POSTED_COL_INDEX = 18  # R열 고정 (원본 시트가 A~Q 17개 컬럼이므로 그 다음 칸)


def pick_unposted_row(ws):
    all_values = ws.get_all_values()
    print(f"[디버그] 시트 전체 행 수: {len(all_values)}")

    # 헤더 행을 자동으로 탐색 (앞에 빈 행이 있어도 안전하게 찾음)
    header_row_idx = None
    for idx, row in enumerate(all_values):
        if row and row[0].strip() == EXPECTED_FIRST_HEADER:
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise RuntimeError(
            f"헤더를 찾지 못함: '{EXPECTED_FIRST_HEADER}'로 시작하는 행이 시트에 없음. "
            f"시트 구조를 확인할 것. 안전을 위해 스크립트가 시트에 아무것도 쓰지 않음."
        )

    header = all_values[header_row_idx]
    sheet_header_row_num = header_row_idx + 1  # 실제 시트상 행 번호(1-based)
    print(f"[디버그] 헤더를 {sheet_header_row_num}행에서 찾음: {header}")

    # 포스팅완료 헤더가 고정 위치(R열)에 없으면 그 위치에만 정확히 기록
    while len(header) < POSTED_COL_INDEX:
        header.append("")
    if header[POSTED_COL_INDEX - 1].strip() != POSTED_COL_NAME:
        ws.update_cell(sheet_header_row_num, POSTED_COL_INDEX, POSTED_COL_NAME)
        header[POSTED_COL_INDEX - 1] = POSTED_COL_NAME
        print(f"[디버그] '{POSTED_COL_NAME}' 헤더를 {sheet_header_row_num}행 {POSTED_COL_INDEX}번째 열(R열)에 기록함")
    else:
        print(f"[디버그] '{POSTED_COL_NAME}' 헤더가 이미 존재함")

    posted_col_idx = POSTED_COL_INDEX - 1  # 0-based

    candidates = []
    data_rows = all_values[header_row_idx + 1:]
    for offset, row_values in enumerate(data_rows):
        sheet_row_num = header_row_idx + 2 + offset  # 실제 시트상 행 번호(1-based)
        posted_val = row_values[posted_col_idx] if posted_col_idx < len(row_values) else ""
        if not posted_val.strip():
            row_dict = dict(zip(header, row_values + [""] * max(0, len(header) - len(row_values))))
            candidates.append((sheet_row_num, row_dict))

    print(f"[디버그] 데이터 행 수: {len(data_rows)}, 포스팅 가능(미완료) 행 수: {len(candidates)}")

    if not candidates:
        raise RuntimeError("포스팅 가능한 상품이 없음 (전부 포스팅완료 상태)")

    chosen_index, chosen_row = random.choice(candidates)
    return chosen_index, chosen_row, header


def mark_posted(ws, row_index, header, post_url):
    value = f"{datetime.now().strftime('%Y-%m-%d %H:%M')} | {post_url}"
    ws.update_cell(row_index, POSTED_COL_INDEX, value)


# ---------- 상품 상세페이지 스크래핑 ----------
def scrape_product_images(detail_url: str):
    res = requests.get(detail_url, headers=UA_HEADERS, timeout=15)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")

    def resolve(src):
        if not src:
            return None
        abs_url = urljoin(detail_url, src)
        if "sample" in abs_url.lower():  # 실제 사진 없는 상품의 기본 샘플 이미지는 제외
            return None
        return abs_url

    # 상품 썸네일 (goods_img 경로) - 상세설명 긴 이미지(dtl_img)는 사용하지 않음
    main_img_tag = soup.select_one(".item_photo_big img")
    if not main_img_tag:
        print("[알림] 썸네일 못씀: 상세페이지에서 .item_photo_big 이미지 태그 자체를 못 찾음")
        main_img = None
    else:
        raw_src = main_img_tag.get("src")
        main_img = resolve(raw_src)
        if main_img is None:
            if raw_src and "sample" in raw_src.lower():
                print(f"[알림] 썸네일 못씀: 이 상품은 실제 사진 없이 기본 샘플 이미지만 등록돼 있음 ({raw_src})")
            else:
                print(f"[알림] 썸네일 못씀: 이미지 src를 찾을 수 없음 (raw_src={raw_src})")

    # 못 찾으면 og:image 메타태그로 대체 시도
    if not main_img:
        og_tag = soup.select_one('meta[property="og:image"]')
        if og_tag and og_tag.get("content"):
            main_img = resolve(og_tag["content"])
            if main_img:
                print(f"[알림] og:image로 대체 확보 성공: {main_img}")
            else:
                print("[알림] 썸네일 못씀: og:image도 없거나 sample 이미지임 → 카테고리 썸네일로 대체 예정")
        else:
            print("[알림] 썸네일 못씀: og:image 메타태그 자체가 없음 → 카테고리 썸네일로 대체 예정")

    title_tag = soup.select_one(".item_detail_tit h3")
    page_title = title_tag.get_text(strip=True) if title_tag else None

    return {
        "main_image": main_img,
        "page_title": page_title,
    }


# ---------- 이미지 다운로드 & repo에 저장 (핫링크 차단 우회) ----------
IMAGES_DIR = "images"


def extract_seq(detail_url: str) -> str:
    if "seq=" in detail_url:
        return detail_url.split("seq=")[-1].split("&")[0]
    return str(random.randint(100000, 999999))


def download_image(url: str, referer: str) -> bytes | None:
    headers = dict(UA_HEADERS)
    headers["Referer"] = referer
    try:
        res = requests.get(url, headers=headers, timeout=15)
        res.raise_for_status()
        print(f"[디버그] 다운로드 성공: {url} ({len(res.content)} bytes, content-type={res.headers.get('Content-Type')})")
        return res.content
    except Exception as e:
        print(f"[경고] 이미지 다운로드 실패: {url} ({e})")
        return None


def download_and_host_product_thumbnail(main_image: str, detail_url: str):
    """상품 썸네일 1장을 다운로드해서 repo의 images/{seq}/ 폴더에 저장하고,
    raw.githubusercontent.com 주소를 반환함 (없으면 None)"""
    if not main_image:
        return None  # scrape_product_images 단계에서 이미 [알림]으로 사유를 남김

    seq = extract_seq(detail_url)
    repo = os.environ.get("GITHUB_REPOSITORY", "")  # 예: rorhkdcns/veluna-auto-post
    branch = "main"

    local_dir = os.path.join(IMAGES_DIR, seq)
    os.makedirs(local_dir, exist_ok=True)

    MAX_WIDTH = 1000  # 가로 1000px 초과시 축소
    JPEG_QUALITY = 80  # 압축 품질 (0~100)

    content = download_image(main_image, referer=detail_url)
    if content is None:
        print(f"[알림] 썸네일 못씀: 이미지 URL은 찾았지만 다운로드가 실패함 ({main_image})")
        return None

    try:
        img = Image.open(io.BytesIO(content))
        img = img.convert("RGB")  # webp/png -> jpg 저장을 위해 변환
        if img.width > MAX_WIDTH:
            new_height = int(img.height * (MAX_WIDTH / img.width))
            img = img.resize((MAX_WIDTH, new_height), Image.LANCZOS)

        local_path = os.path.join(local_dir, "thumb.jpg")
        img.save(local_path, "JPEG", quality=JPEG_QUALITY, optimize=True)
    except Exception as e:
        print(f"[경고] 이미지 처리 실패, 원본 그대로 저장: {main_image} ({e})")
        ext = main_image.split(".")[-1].split("?")[0]
        if len(ext) > 5:
            ext = "webp"
        local_path = os.path.join(local_dir, f"thumb.{ext}")
        with open(local_path, "wb") as f:
            f.write(content)

    raw_url = f"https://raw.githubusercontent.com/{repo}/{branch}/{local_path}"
    print(f"[디버그] 상품 썸네일 저장 완료: {raw_url}")
    return raw_url


# ---------- Gemini로 포스팅 본문 생성 ----------
ACCENT_COLOR = "#E75480"  # 키워드 강조색 (핑크톤)


def generate_post_content(product_name: str, category_label: str, category_raw: str,
                           price: int, detail_url: str) -> dict:
    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

    prompt = f"""
너는 성인용품 전문 쇼핑몰 '벨루나(velunamall.com)'의 제휴 블로그 작성자야.
아래 상품 정보를 바탕으로 성인 인증된 독자를 대상으로 하는 블로그 포스팅을 작성해.

[상품 정보]
- 상품명: {product_name}
- 분류: {category_label}용품 (세부 카테고리: {category_raw})
- 가격: {price:,}원
- 구매링크: {detail_url}

[작성 규칙]
- 제목은 반드시 "[성인용품 벨루나]"로 시작할 것 (그 뒤에 상품 특징을 살린 매력적인 제목 이어붙이기)
- 제목에 "솔직후기", "내돈내산" 같은 후기 프레이밍 문구는 절대 넣지 말 것 (정보성 제목으로)
- 제품명은 노골적으로 그대로 사용해도 됨 (성인 인증된 사이트이므로)
- 과도하게 선정적인 묘사보다는 제품 특징, 소재, 사용 편의성, 추천 대상 위주로 정보성 있게 작성
- 글 전체에 구글 검색품질 평가 기준인 E-E-A-T(경험/전문성/권위성/신뢰성)가 자연스럽게 드러나도록 작성할 것:
  실제 사용 경험을 연상시키는 구체적 디테일, 소재/원리에 대한 전문적 설명, 근거 있는 주의사항이나 관리법을 포함
- 글의 설득 흐름은 PASONA 법칙(문제 제기 → 문제 공감·심화 → 해결책 제시 → 제안 → 대상 좁히기 → 행동 유도)을
  자연스럽게 따르되, 이런 법칙 이름이나 단계 이름은 절대 본문에 표기하지 말고 매끄러운 글로만 녹여낼 것
- 각 섹션마다 핵심 키워드(소재명, 기능명, 특징 등) 3~5개를 골라 <strong style="color:{ACCENT_COLOR};">키워드</strong> 형태로 볼드+컬러 강조할 것
- 본문 텍스트 안에는 절대 <a> 링크 태그를 넣지 말 것 (구매 링크는 스크립트가 별도로 카드 형태로 삽입함)
- 아래 포맷을 반드시 지킬 것:
- 문체는 딱딱한 정보전달체가 아니라, 친한 사람이 조곤조곤 이야기해주듯 부드럽고 다정한 말투(해요체)로 쓸 것.
  단정적이고 건조한 "~합니다/~됩니다" 나열보다는, 공감하는 말투와 자연스러운 구어체 표현을 섞어서 쓸 것
- "도입부", "본문", "결론" 같은 구조를 그대로 드러내는 라벨/소제목은 절대 쓰지 말 것.
  인트로 문단은 소제목 없이 바로 시작하고, 각 섹션은 실제 주제를 담은 소제목(예: "1. 관계에서 느끼는 미묘한 권태와...")만 h2로 쓰고
  그 소제목들을 다시 "본문"이라는 상위 라벨로 묶지 말 것. 마지막 문단도 "결론"이라는 라벨 없이 소제목 없이 자연스럽게 마무리할 것
- 아래 내용 순서를 지키되, 위에서 말한 라벨은 쓰지 말 것:
  1. 제목 (매력적이고 검색엔진 친화적으로)
  2. 인트로 문단 (소제목 없이, 2~3문장)
  3. 목차 (섹션 제목 리스트, ul 또는 ol 태그로)
  4. 섹션별 본문 (각 섹션 제목을 h2로, 그 안에 필요하면 h3 소제목)
  5. 요약 박스가 될 내용 (핵심 포인트 3~4개, 불릿 리스트로. 박스 스타일은 스크립트가 감싸줄 예정이니 그냥 ul/li로만 작성)
  6. FAQ (질문 3개 + 답변, h3나 strong으로 질문 표시)
  7. 마무리 문단 (소제목 없이, 구매를 자연스럽게 유도하되 링크는 절대 넣지 말 것 - 카드가 별도로 붙음)

출력은 JSON 형식으로만 응답해. 다른 텍스트/설명/마크다운 코드블록 없이 아래 스키마 그대로:
{{
  "title": "포스팅 제목",
  "html_body": "완성된 HTML 본문 전체 (h2/h3/p/ul/li/strong 태그 사용, 이미지·링크 태그는 넣지 말 것 - 스크립트가 별도로 삽입함)"
}}
"""

    response = client.models.generate_content(
        model="gemini-flash-latest",
        contents=prompt,
        config={"response_mime_type": "application/json"},
    )
    text = response.text.strip()
    text = text.replace("```json", "").replace("```", "").strip()

    # 앞뒤에 잡텍스트(예: ***, 설명 문구 등)가 붙어도 첫 '{' ~ 마지막 '}' 사이만 추출
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"[경고] JSON 파싱 실패, 응답 원문 일부: {text[:500]}")
        raise


# ---------- 구매 유도 이미지 카드 ----------
def build_purchase_card_html(product_name: str, price: int, card_image_url: str, detail_url: str) -> str:
    """텍스트 링크 대신, 이미지+상품명+버튼이 하나의 카드로 묶여
    전체가 클릭 가능한 구매 유도 카드를 만듦"""
    image_html = (
        f'<img src="{card_image_url}" alt="{product_name}" '
        f'style="width:100%;max-height:280px;object-fit:cover;display:block;'
        f'border-top-left-radius:14px;border-top-right-radius:14px;">'
        if card_image_url else ""
    )

    return f"""
<a href="{detail_url}" target="_blank" rel="nofollow"
   style="text-decoration:none; display:block; max-width:480px; margin:40px auto;
          border-radius:14px; overflow:hidden; background:#1c1c1c;
          box-shadow:0 4px 16px rgba(0,0,0,0.25); border:1px solid #333;">
  {image_html}
  <div style="padding:22px 24px; clear:both;">
    <p style="margin:0 0 8px 0; color:#ff9ec4; font-size:14px; letter-spacing:1px; font-weight:600;">VELUNA MALL</p>
    <p style="margin:0 0 14px 0; color:#ffffff; font-size:20px; font-weight:700; line-height:1.4;">{product_name}</p>
    <p style="margin:0 0 20px 0; color:#f2f2f2; font-size:18px; font-weight:600;">{price:,}원</p>
    <div style="text-align:center; padding:16px 0; border-radius:10px;
                background:linear-gradient(135deg,#E75480,#ff8fab);
                color:#ffffff; font-weight:700; font-size:18px;">
      지금 구매하러 가기 →
    </div>
  </div>
</a>
""".strip()


# ---------- 제목 정규화 ----------
TITLE_PREFIX = "[성인용품 벨루나]"
BANNED_TITLE_PATTERNS = [r"솔직\s*후기", r"내돈\s*내산", r"리얼\s*후기", r"실사용\s*후기"]


def normalize_title(title: str) -> str:
    t = title.strip()

    # 금지 문구 제거 (띄어쓰기 변형 포함)
    for pattern in BANNED_TITLE_PATTERNS:
        t = re.sub(pattern, "", t)

    # 기존에 다른 형태의 대괄호 접두어가 붙어있으면 제거 후 통일된 접두어로 교체
    t = re.sub(r"^\[[^\]]*\]\s*", "", t).strip()
    t = re.sub(r"\s{2,}", " ", t).strip(" ,.-")

    return f"{TITLE_PREFIX} {t}"


# ---------- 목차 앵커 링크 처리 ----------
def add_toc_links(html_body: str) -> str:
    """본문의 h2 소제목마다 id를 붙이고, 그 앞에 나오는 목차 리스트(ul)를
    클릭하면 해당 소제목으로 이동하는 링크로 바꿔줌"""
    try:
        soup = BeautifulSoup(html_body, "html.parser")
        h2_tags = soup.find_all("h2")
        if not h2_tags:
            return html_body

        for i, h2 in enumerate(h2_tags, start=1):
            h2["id"] = f"section-{i}"

            # 소제목 텍스트에서 앞자리 번호(예: "1. ") 제거하고, 대신 동그란 번호 뱃지로 표시
            raw_text = h2.get_text(strip=True)
            m = re.match(r"^\d+[.\)]\s*(.+)$", raw_text)
            title_text = m.group(1) if m else raw_text

            h2.clear()
            h2["style"] = "display:flex; align-items:center; gap:10px;"
            badge = soup.new_tag("span", style=(
                f"display:inline-flex; align-items:center; justify-content:center; "
                f"min-width:30px; height:30px; border-radius:50%; background:{ACCENT_COLOR}; "
                f"color:#fff; font-size:15px; font-weight:700; flex-shrink:0;"
            ))
            badge.string = str(i)
            h2.append(badge)
            h2.append(title_text)

        # 첫 h2보다 앞에 나오는 리스트(ul 또는 ol) 중 가장 가까운 것 = 목차일 가능성이 높음
        first_h2 = h2_tags[0]
        toc_list = None
        for el in first_h2.find_all_previous(["ul", "ol"]):
            toc_list = el  # find_all_previous는 가까운 순이라 첫번째가 가장 가까움
            break

        if toc_list is not None:
            items = toc_list.find_all("li", recursive=False)
            for i, li in enumerate(items, start=1):
                if i > len(h2_tags):
                    break
                text = li.get_text(strip=True)
                a_tag = soup.new_tag("a", href=f"#section-{i}", style="color:inherit; text-decoration:none;")
                a_tag.string = text
                li.clear()
                li.append(a_tag)

            # 목차 전체를 테두리 박스로 감싸기
            toc_list["style"] = (
                "border:1px solid #e0a8bd; border-radius:10px; padding:16px 20px 16px 36px; "
                "background:#fff5f8; margin:20px 0; line-height:1.9;"
            )

        return str(soup)
    except Exception as e:
        print(f"[경고] 목차 링크 처리 실패, 원본 그대로 사용: {e}")
        return html_body


# ---------- Blogger 업로드 ----------
def get_blogger_service():
    token_b64 = os.environ["BLOGGER_TOKEN_PICKLE_B64"]
    creds = pickle.loads(base64.b64decode(token_b64))
    return build("blogger", "v3", credentials=creds)


def post_to_blogger(title: str, content_html: str) -> str:
    service = get_blogger_service()
    blog_id = os.environ["BLOGGER_BLOG_ID"]

    body = {
        "kind": "blogger#post",
        "title": title,
        "content": content_html,
    }

    result = service.posts().insert(blogId=blog_id, body=body, isDraft=False).execute()
    return result.get("url")


# ---------- 메인 ----------
def main():
    ws = get_sheet()

    row_index, row, header = pick_unposted_row(ws)

    product_name = row.get("본사상품명(변경불가)") or row.get("내상품명") or "상품"
    category_raw = row.get("카테고리(변경불가)", "")
    detail_url = row.get("제품상세URL(변경불가)")
    price = int(row.get("일반가(변경불가)") or 0)

    if not detail_url:
        raise RuntimeError(f"제품상세URL 없음: {row}")

    category_label = classify(category_raw)

    scraped = scrape_product_images(detail_url)

    product_thumb_url = download_and_host_product_thumbnail(scraped["main_image"], detail_url)
    category_thumb_url = get_category_thumbnail_url(category_label)

    if product_thumb_url:
        print(f"[결과] 상품 자체 썸네일 사용함: {product_thumb_url}")
    else:
        print(f"[결과] 상품 자체 썸네일 못씀 → 카테고리({category_label}) 고정 썸네일로 대체함: {category_thumb_url}")

    content = generate_post_content(
        product_name=product_name,
        category_label=category_label,
        category_raw=category_raw,
        price=price,
        detail_url=detail_url,
    )
    content["html_body"] = add_toc_links(content["html_body"])
    content["title"] = normalize_title(content["title"])

    category_thumb_html = f'<p style="text-align:center;margin:0 0 24px 0;"><img src="{category_thumb_url}" alt="{category_label}용품" style="max-width:100%;"></p>'

    product_thumb_html = (
        f'<p style="text-align:center;margin:24px 0;"><img src="{product_thumb_url}" alt="{product_name}" style="max-width:100%;"></p>'
        if product_thumb_url else ""
    )

    card_image_url = product_thumb_url or category_thumb_url
    purchase_card_html = build_purchase_card_html(product_name, price, card_image_url, detail_url)

    # 최종 포스팅 HTML: 카테고리 썸네일 -> 본문 -> 실제 상품 썸네일 -> 구매 카드
    final_html = f"""
{category_thumb_html}
{content['html_body']}
{product_thumb_html}
{purchase_card_html}
""".strip()

    post_url = post_to_blogger(content["title"], final_html)
    mark_posted(ws, row_index, header, post_url)

    print(f"포스팅 완료: {content['title']}")
    print(f"URL: {post_url}")


if __name__ == "__main__":
    main()
