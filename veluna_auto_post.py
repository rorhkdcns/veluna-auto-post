# -*- coding: utf-8 -*-
"""
벨루나몰(velunamall.com) 상품 -> 블로그 자동 포스팅 스크립트 (Imgur 적용 + 중간 구매카드 삽입)

흐름:
1. Google Sheet에서 '포스팅완료'가 비어있는 상품 중 랜덤 1개 선택
2. 상품상세URL 스크래핑 -> 실제 상품 이미지(썸네일) 추출
3. 카테고리 -> 남성/여성/커플 분류 (category_map.py)
4. 이미지를 Imgur API로 전송하여 새로운 URL 발급
5. Gemini API로 포스팅 본문 생성
6. Blogger API로 포스팅 업로드 (중간 및 하단 구매카드 포함)
7. 포스팅 완료된 행에 '포스팅완료' 표시
"""

import os
import json
import random
import base64
import pickle
import re
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
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
    # 탭 이름 에러 방지: 첫 번째 탭 강제 사용
    return sh.get_worksheet(0)

EXPECTED_FIRST_HEADER = "인덱스(변경불가)"

def pick_unposted_row(ws):
    all_values = ws.get_all_values()
    print(f"[디버그] 시트 전체 행 수: {len(all_values)}")

    header_row_idx = None
    for idx, row in enumerate(all_values):
        if row and row[0].strip() == EXPECTED_FIRST_HEADER:
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise RuntimeError(
            f"헤더를 찾지 못함: '{EXPECTED_FIRST_HEADER}'로 시작하는 행이 시트에 없음."
        )

    header = all_values[header_row_idx]
    sheet_header_row_num = header_row_idx + 1
    print(f"[디버그] 헤더를 {sheet_header_row_num}행에서 찾음")

    if POSTED_COL_NAME in header:
        posted_col_idx = header.index(POSTED_COL_NAME)
        posted_col_num = posted_col_idx + 1
    else:
        header.append(POSTED_COL_NAME)
        posted_col_idx = len(header) - 1
        posted_col_num = len(header)
        ws.update_cell(sheet_header_row_num, posted_col_num, POSTED_COL_NAME)
        print(f"[디버그] '{POSTED_COL_NAME}' 헤더를 {posted_col_num}번째 열에 새로 추가함")

    candidates = []
    data_rows = all_values[header_row_idx + 1:]
    for offset, row_values in enumerate(data_rows):
        sheet_row_num = header_row_idx + 2 + offset
        posted_val = row_values[posted_col_idx] if posted_col_idx < len(row_values) else ""
        if not posted_val.strip():
            row_dict = dict(zip(header, row_values + [""] * max(0, len(header) - len(row_values))))
            candidates.append((sheet_row_num, row_dict))

    print(f"[디버그] 데이터 행 수: {len(data_rows)}, 포스팅 가능(미완료) 행 수: {len(candidates)}")

    if not candidates:
        raise RuntimeError("포스팅 가능한 상품이 없음 (전부 포스팅완료 상태)")

    chosen_index, chosen_row = random.choice(candidates)
    return chosen_index, chosen_row, posted_col_num

def mark_posted(ws, row_index, posted_col_num, post_url):
    value = f"{datetime.now().strftime('%Y-%m-%d %H:%M')} | {post_url}"
    ws.update_cell(row_index, posted_col_num, value)

# ---------- 상품 상세페이지 스크래핑 ----------
def scrape_product_images(detail_url: str):
    res = requests.get(detail_url, headers=UA_HEADERS, timeout=15)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")

    def resolve(src):
        if not src: return None
        abs_url = urljoin(detail_url, src)
        if "sample" in abs_url.lower(): return None
        return abs_url

    main_img_tag = soup.select_one(".item_photo_big img")
    main_img = None
    if main_img_tag:
        main_img = resolve(main_img_tag.get("src"))

    if not main_img:
        og_tag = soup.select_one('meta[property="og:image"]')
        if og_tag and og_tag.get("content"):
            main_img = resolve(og_tag["content"])

    title_tag = soup.select_one(".item_detail_tit h3")
    page_title = title_tag.get_text(strip=True) if title_tag else None

    return {
        "main_image": main_img,
        "page_title": page_title,
    }

# ---------- Imgur API를 활용한 이미지 업로드 ----------
def upload_thumbnail_to_imgur(image_url: str, referer: str) -> str:
    if not image_url:
        return None

    headers = dict(UA_HEADERS)
    headers["Referer"] = referer
    try:
        res = requests.get(image_url, headers=headers, timeout=15)
        res.raise_for_status()
    except Exception as e:
        print(f"[경고] 쇼핑몰 이미지 임시 다운로드 실패: {image_url} ({e})")
        return None

    imgur_client_id = os.environ.get("IMGUR_CLIENT_ID")
    if not imgur_client_id:
        raise RuntimeError("IMGUR_CLIENT_ID 환경변수가 설정되지 않았습니다.")

    imgur_headers = {"Authorization": f"Client-ID {imgur_client_id}"}
    try:
        upload_res = requests.post(
            "https://api.imgur.com/3/image",
            headers=imgur_headers,
            files={"image": res.content},
            timeout=20
        )
        upload_res.raise_for_status()
        raw_url = upload_res.json()["data"]["link"]
        print(f"[디버그] Imgur 업로드 성공 (핫링크 우회 완료): {raw_url}")
        return raw_url
    except Exception as e:
        print(f"[경고] Imgur 업로드 실패: {e}")
        return None

# ---------- Gemini로 포스팅 본문 생성 ----------
ACCENT_COLOR = "#E75480"

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
- 글 전체에 구글 검색품질 평가 기준인 E-E-A-T(경험/전문성/권위성/신뢰성)가 자연스럽게 드러나도록 작성할 것
- 글의 설득 흐름은 PASONA 법칙을 자연스럽게 따르되, 이런 법칙 이름이나 단계 이름은 절대 본문에 표기하지 말 것
- 각 섹션마다 핵심 키워드(소재명, 기능명, 특징 등) 3~5개를 골라 <strong style="color:{ACCENT_COLOR};">키워드</strong> 형태로 볼드+컬러 강조할 것
- 본문 텍스트 안에는 절대 <a> 링크 태그를 넣지 말 것 (구매 링크는 스크립트가 별도로 카드 형태로 삽입함)
- 문체는 친한 사람이 조곤조곤 이야기해주듯 부드럽고 다정한 말투(해요체)로 쓸 것.
- "도입부", "본문", "결론" 같은 구조를 그대로 드러내는 라벨/소제목은 절대 쓰지 말 것.
  인트로 문단은 소제목 없이 바로 시작하고, 각 섹션은 실제 주제를 담은 소제목만 h2로 쓸 것.
- 아래 내용 순서를 지키되, 위에서 말한 라벨은 쓰지 말 것:
  1. 제목 (매력적이고 검색엔진 친화적으로)
  2. 인트로 문단 (소제목 없이, 2~3문장)
  3. 목차 (섹션 제목 리스트, ul 또는 ol 태그로)
  4. 섹션별 본문 (각 섹션 제목을 h2로, 그 안에 필요하면 h3 소제목)
  5. 요약 박스 (핵심 포인트 3~4개, 불릿 리스트로. 반드시 아래처럼 감싸서 출력할 것:
     <div class="summary-box"><ul><li>포인트1</li><li>포인트2</li>...</ul></div>)
  6. FAQ (질문 3개 + 답변, h3나 strong으로 질문 표시)
  7. 마무리 문단 (소제목 없이, 구매를 자연스럽게 유도하되 링크는 절대 넣지 말 것)

출력 형식은 JSON이 아니라 아래 구분자 형식으로만 출력할 것:

[TITLE]
여기에 포스팅 제목만 한 줄로

[BODY]
여기에 완성된 HTML 본문 전체 (이미지·링크 태그 제외)
"""
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )
    text = response.text.strip()
    text = text.replace("```html", "").replace("```json", "").replace("```", "").strip()

    if "[BODY]" not in text:
        raise RuntimeError("Gemini 응답 형식이 예상과 다름 ([BODY] 구분자 없음)")

    title_part, body_part = text.split("[BODY]", 1)
    title = title_part.replace("[TITLE]", "").strip()
    html_body = body_part.strip()

    return {"title": title, "html_body": html_body}

# ---------- 구매 유도 이미지 카드 ----------
def build_purchase_card_html(product_name: str, price: int, card_image_url: str, detail_url: str) -> str:
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
    for pattern in BANNED_TITLE_PATTERNS:
        t = re.sub(pattern, "", t)
    t = re.sub(r"^\[[^\]]*\]\s*", "", t).strip()
    t = re.sub(r"\s{2,}", " ", t).strip(" ,.-")
    return f"{TITLE_PREFIX} {t}"

# ---------- HTML 본문 후처리 (요약 박스, 목차, 중간 카드 삽입) ----------
def style_summary_box(html_body: str) -> str:
    try:
        soup = BeautifulSoup(html_body, "html.parser")
        box = soup.find("div", class_="summary-box")
        if box is None: return html_body

        box["style"] = (
            "border:1px solid #f0b8cc; border-left:5px solid #E75480; border-radius:10px; "
            "background:#fff5f8; padding:20px 24px; margin:28px 0;"
        )
        heading = soup.new_tag("p", style="margin:0 0 12px 0; font-weight:700; font-size:16px; color:#E75480;")
        heading.string = "💡 이것만은 꼭 기억하세요"
        box.insert(0, heading)
        return str(soup)
    except:
        return html_body

def add_toc_links(html_body: str) -> str:
    try:
        soup = BeautifulSoup(html_body, "html.parser")
        h2_tags = soup.find_all("h2")
        if not h2_tags: return html_body

        for i, h2 in enumerate(h2_tags, start=1):
            h2["id"] = f"section-{i}"
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

        first_h2 = h2_tags[0]
        toc_list = None
        for el in first_h2.find_all_previous(["ul", "ol"]):
            toc_list = el
            break

        if toc_list is not None:
            items = toc_list.find_all("li", recursive=False)
            for i, li in enumerate(items, start=1):
                if i > len(h2_tags): break
                text = li.get_text(strip=True)
                a_tag = soup.new_tag("a", href=f"#section-{i}", style="color:inherit; text-decoration:none;")
                a_tag.string = text
                li.clear()
                li.append(a_tag)

            toc_list["style"] = (
                "border:1px solid #e0a8bd; border-radius:10px; padding:16px 20px 16px 36px; "
                "background:#fff5f8; margin:20px 0; line-height:1.9;"
            )
        return str(soup)
    except:
        return html_body

# 💡 새롭게 추가된 기능: 본문 중간에 구매 카드 삽입
def insert_middle_purchase_card(html_body: str, card_html: str) -> str:
    """본문의 섹션(h2 태그) 개수를 파악하여 정확히 중간 위치에 구매 카드를 추가합니다."""
    try:
        soup = BeautifulSoup(html_body, "html.parser")
        h2_tags = soup.find_all("h2")
        
        # 섹션이 2개 이상일 때만 본문 중간에 삽입
        if len(h2_tags) >= 2:
            mid_idx = len(h2_tags) // 2
            target_h2 = h2_tags[mid_idx]
            
            card_soup = BeautifulSoup(card_html, "html.parser")
            target_h2.insert_before(card_soup)
            return str(soup)
        
        return html_body
    except Exception as e:
        print(f"[경고] 중간 구매 카드 삽입 실패: {e}")
        return html_body

# ---------- Blogger 업로드 ----------
def get_blogger_service():
    token_b64 = os.environ["BLOGGER_TOKEN_PICKLE_B64"]
    creds = pickle.loads(base64.b64decode(token_b64))
    return build("blogger", "v3", credentials=creds)

def post_to_blogger(title: str, content_html: str) -> str:
    service = get_blogger_service()
    blog_id = os.environ["BLOGGER_BLOG_ID"]

    body = {"kind": "blogger#post", "title": title, "content": content_html}
    result = service.posts().insert(blogId=blog_id, body=body, isDraft=False).execute()
    return result.get("url")

# ---------- 메인 ----------
def main():
    ws = get_sheet()

    row_index, row, posted_col_num = pick_unposted_row(ws)

    product_name = row.get("본사상품명(변경불가)") or row.get("내상품명") or "상품"
    category_raw = row.get("카테고리(변경불가)", "")
    detail_url = row.get("제품상세URL(변경불가)")
    
    price_str = str(row.get("일반가(변경불가)", "0"))
    price = int(re.sub(r'[^0-9]', '', price_str) or 0)

    if not detail_url:
        raise RuntimeError(f"제품상세URL 없음: {row}")

    category_label = classify(category_raw)
    scraped = scrape_product_images(detail_url)

    product_thumb_url = upload_thumbnail_to_imgur(scraped["main_image"], detail_url)
    category_thumb_url = get_category_thumbnail_url(category_label)

    content = generate_post_content(
        product_name=product_name,
        category_label=category_label,
        category_raw=category_raw,
        price=price,
        detail_url=detail_url,
    )
    
    card_image_url = product_thumb_url or category_thumb_url
    purchase_card_html = build_purchase_card_html(product_name, price, card_image_url, detail_url)

    # 💡 HTML 후처리 단계 (순서: 요약박스 -> 목차 -> ❗중간 카드 삽입❗)
    content["html_body"] = style_summary_box(content["html_body"])
    content["html_body"] = add_toc_links(content["html_body"])
    content["html_body"] = insert_middle_purchase_card(content["html_body"], purchase_card_html)
    content["title"] = normalize_title(content["title"])

    category_thumb_html = f'<p style="text-align:center;margin:0 0 24px 0;"><img src="{category_thumb_url}" alt="{category_label}용품" style="max-width:100%;"></p>'
    product_thumb_html = (
        f'<p style="text-align:center;margin:24px 0;"><img src="{product_thumb_url}" alt="{product_name}" style="max-width:100%;"></p>'
        if product_thumb_url else ""
    )

    # 최종 병합: (하단에도 기존처럼 카드가 들어갑니다)
    final_html = f"""
{category_thumb_html}
{content['html_body']}
{product_thumb_html}
{purchase_card_html}
""".strip()

    post_url = post_to_blogger(content["title"], final_html)
    mark_posted(ws, row_index, posted_col_num, post_url)

    print(f"포스팅 완료: {content['title']}")
    print(f"URL: {post_url}")

if __name__ == "__main__":
    main()
