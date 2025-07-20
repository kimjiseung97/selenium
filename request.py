import requests
import openpyxl
import time
import httpx

def get_reviews(product_id):
    all_reviews = []
    page = 1

    headers = {
        # 최신 Chrome User-Agent로 업데이트 (주기적으로 변경될 수 있으니 최신 값 확인 필요)
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'Accept': 'application/json, text/plain, */*',
        'Connection': 'keep-alive',
    }


    with httpx.Client(http2=True, headers=headers, timeout=10) as client:
        while True:
            url = f"https://www.coupang.com/next-api/review?productId={product_id}&page={page}&size=10&sortBy=ORDER_SCORE_ASC&ratingSummary=true&ratings=&market="
            try:
                resp = client.get(url)
                if resp.status_code != 200:
                    print(f"❌ HTTP {resp.status_code} 실패")
                    break

                data = resp.json()
                contents = data.get("rData", {}).get("contents", [])
                if not contents:
                    print(f"📄 더 이상 리뷰 없음 (page {page})")
                    break

                for review in contents:
                    all_reviews.append({
                        "작성자": review.get("authorName", ""),
                        "작성일": review.get("createdAt", ""),
                        "평점": review.get("rating", ""),
                        "리뷰내용": review.get("reviewContent", "").replace('\n', ' ')
                    })

                print(f"✅ {page} 페이지 수집 완료 (누적: {len(all_reviews)}개)")
                total_page = data.get("rData", {}).get("paging", {}).get("totalPage", 1)
                if page >= total_page:
                    break
                page += 1
                time.sleep(0.5)
            except Exception as e:
                print(f"❗ 예외 발생 (page {page}): {e}")
                break
    return all_reviews    

    


def save_to_excel(reviews, filename="coupang_reviews.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["작성자", "작성일", "평점", "리뷰내용"])

    for r in reviews:
        ws.append([r["작성자"], r["작성일"], r["평점"], r["리뷰내용"]])

    wb.save(filename)
    print(f"📁 저장 완료: {filename} (총 {len(reviews)}개 리뷰)")


if __name__ == "__main__":
    product_id = "8243821628"
    print("🚀 리뷰 수집 시작")
    reviews = get_reviews(product_id)
    save_to_excel(reviews)
