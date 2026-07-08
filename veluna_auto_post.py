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
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from google import genai

from category_map import classify

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


def pick_unposted_row(ws):
    all_values = ws.get_all_values()
    header = all_values[0]

    # 빈 헤더 셀 제거 (뒤쪽 빈 열들 때문에 발생하는 중복 에러 방지)
    last_col = len(header)
    while last_col > 0 and not header[last_col - 1].strip():
        last_col -= 1
    header = header[:last_col]

    if POSTED_COL_NAME not in header:
        # 없으면 마지막 열에 헤더 추가
        ws.update_cell(1, len(header) + 1, POSTED_COL_NAME)
        header.append(POSTED_COL_NAME)
        last_col += 1

    posted_col_idx = header.index(POSTED_COL_NAME)

    candidates = []
    for i, row_values in enumerate(all_values[1:], start=2):  # 2행부터 실제 데이터
        row_values = row_values[:last_col] + [""] * max(0, last_col - len(row_values))
        posted_val = row_values[posted_col_idx] if posted_col_idx < len(row_values) else ""
        if not posted_val.strip():
            row_dict = dict(zip(header, row_values))
            candidates.append((i, row_dict))

    if not candidates:
        raise RuntimeError("포스팅 가능한 상품이 없음 (전부 포스팅완료 상태)")

    chosen_index, chosen_row = random.choice(candidates)
    return chosen_index, chosen_row, header


def mark_posted(ws, row_index, header, post_url):
    col_idx = header.index(POSTED_COL_NAME) + 1
    value = f"{datetime.now().strftime('%Y-%m-%d %H:%M')} | {post_url}"
    ws.update_cell(row_index, col_idx, value)


# ---------- 상품 상세페이지 스크래핑 ----------
def scrape_product_images(detail_url: str):
    res = requests.get(detail_url, headers=UA_HEADERS, timeout=15)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")

    # 메인(대표) 이미지
    main_img_tag = soup.select_one(".item_photo_big img")
    main_img = main_img_tag["src"] if main_img_tag else None

    # 상세설명 이미지들 (실제 상품 사진 나열)
    detail_imgs = [img["src"] for img in soup.select(".viewimg img") if img.get("src")]

    # 실제 상품명 (엑셀/시트 값과 다를 수 있어 검증용으로 같이 반환)
    title_tag = soup.select_one(".item_detail_tit h3")
    page_title = title_tag.get_text(strip=True) if title_tag else None

    return {
        "main_image": main_img,
        "detail_images": detail_imgs,
        "page_title": page_title,
    }


# ---------- Gemini로 포스팅 본문 생성 ----------
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
- 제품명은 노골적으로 그대로 사용해도 됨 (성인 인증된 사이트이므로)
- 과도하게 선정적인 묘사보다는 제품 특징, 소재, 사용 편의성, 추천 대상 위주로 정보성 있게 작성
- 아래 포맷을 반드시 지킬 것:
  1. 제목 (매력적이고 검색엔진 친화적으로)
  2. 도입부 (2~3문장)
  3. 목차 (섹션 제목 리스트)
  4. 본문 (섹션별로 나눠서 작성: 제품 특징, 이런 분께 추천, 사용 팁 등)
  5. 요약 박스 (핵심 포인트 3~4개, 불릿)
  6. FAQ (질문 3개 + 답변)
  7. 결론 (구매링크 자연스럽게 언급)

출력은 JSON 형식으로만 응답해. 다른 텍스트/설명/마크다운 코드블록 없이 아래 스키마 그대로:
{{
  "title": "포스팅 제목",
  "html_body": "완성된 HTML 본문 전체 (h2/h3/p/ul/li 태그 사용, 이미지 태그는 넣지 말 것 - 이미지는 스크립트가 별도로 삽입함)"
}}
"""

    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=prompt,
    )
    text = response.text.strip()
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)


# ---------- 이미지 HTML 블록 생성 ----------
def build_image_html(main_image: str, detail_images: list) -> tuple[str, str]:
    """썸네일 이미지 HTML, 상세이미지 HTML 두 개를 나눠서 반환"""
    thumb_html = f'<p style="text-align:center;"><img src="{main_image}" alt="상품 대표 이미지" style="max-width:100%;"></p>' if main_image else ""

    detail_html_parts = []
    for src in detail_images:
        detail_html_parts.append(
            f'<p style="text-align:center;margin:0;"><img src="{src}" style="max-width:100%;"></p>'
        )
    detail_html = "\n".join(detail_html_parts)

    return thumb_html, detail_html


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
    main_image = scraped["main_image"]
    detail_images = scraped["detail_images"]

    content = generate_post_content(
        product_name=product_name,
        category_label=category_label,
        category_raw=category_raw,
        price=price,
        detail_url=detail_url,
    )

    thumb_html, detail_img_html = build_image_html(main_image, detail_images)

    # 최종 포스팅 HTML: 썸네일 -> 본문 -> 상세이미지 -> 구매링크
    final_html = f"""
{thumb_html}
{content['html_body']}
<h3>상품 상세 이미지</h3>
{detail_img_html}
<p style="text-align:center;margin-top:20px;">
  <a href="{detail_url}" target="_blank" rel="nofollow">👉 상품 자세히 보러가기</a>
</p>
""".strip()

    post_url = post_to_blogger(content["title"], final_html)
    mark_posted(ws, row_index, header, post_url)

    print(f"포스팅 완료: {content['title']}")
    print(f"URL: {post_url}")


if __name__ == "__main__":
    main()
