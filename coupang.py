import math
import time
import re
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, NoSuchElementException
from openpyxl import Workbook
from selenium.webdriver.chrome.service import Service as ChromeService


def setup_driver():
    options = uc.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--start-maximized')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    driver = uc.Chrome(options=options)
    return driver

def save_to_excel(reviews, filename="coupang_reviews.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.append(["작성자", "작성일", "평점", "리뷰 내용"])
    for r in reviews:
        ws.append([r["작성자"], r["작성일"], r["평점"], r["리뷰내용"]])
    wb.save(filename)
    
def get_review_totalcount(driver):        
        review_count_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "상품평")]'))
        )
        text = review_count_element.text
        match = re.search(r'([\d,]+)', text)
        total_reviews = int(match.group(1).replace(',', '')) if match else 0
        return total_reviews
    
def click_next_page(driver, current_page):
    try:
        old_usernames = [
            e.text for e in driver.find_elements(By.CSS_SELECTOR, 'span.sdp-review__article__list__info__user__name')
        ]
    except:
        old_usernames = []
    try:
        next_page_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, f'button.js_reviewArticlePageBtn[data-page="{current_page + 1}"]'))
        )
        # next_page_btn.click()
        driver.execute_script("arguments[0].click();", next_page_btn)
        current_page += 1
        time.sleep(1.5)

        WebDriverWait(driver, 5).until(
            lambda d: any(
                e.text not in old_usernames for e in d.find_elements(By.CSS_SELECTOR, 'span.sdp-review__article__list__info__user__name')
            )
        )
    except:
        try:
            # 다음 묶음으로 넘기기 (▶ 버튼)
            next_group_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.sdp-review__article__page__next'))
            )
            driver.execute_script("arguments[0].click();", next_group_btn)
            time.sleep(1.5)

            next_page_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, f'button.js_reviewArticlePageBtn[data-page="{current_page + 1}"]'))
            )
            driver.execute_script("arguments[0].click();", next_page_btn)
            current_page += 1
            time.sleep(1.5)

            WebDriverWait(driver, 5).until(
                lambda d: any(
                    e.text not in old_usernames for e in d.find_elements(By.CSS_SELECTOR, 'span.sdp-review__article__list__info__user__name')
                )
            )
        except:
            current_page += 1  # 실패해도 무한루프 방지용 강제 증가
    return current_page
        

def crawl_reviews(url,driver):
    driver.get(url)
    time.sleep(2)
    
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "상품평")]'))
        ).click()
        time.sleep(2)
    except UnexpectedAlertPresentException:
        try:
            alert = driver.switch_to.alert
            print("❌ 쿠팡에서 차단되었습니다. Alert 메시지:", alert.text)
            alert.accept()
        except:
            pass
        driver.quit()
        return []
    except Exception as e:
        print("❌ 리뷰 탭 진입 실패:", e)
        driver.quit()
        return []

    reviews = []
    try:
        total_review_count = get_review_totalcount(driver)
        total_pages = math.ceil(total_review_count / 10)
        current_page = 1
        for _ in range(1, total_pages + 1):
            try:
                articles = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'article.sdp-review__article__list'))
                )
            except:
                continue

            for art in articles:
                try:
                    username = art.find_element(By.CSS_SELECTOR, 'span.sdp-review__article__list__info__user__name').text
                except NoSuchElementException:
                    username = ""

                try:
                    date = art.find_element(By.CSS_SELECTOR, 'div.sdp-review__article__list__info__product-info__reg-date').text
                except NoSuchElementException:
                    date = ""

                try:
                    content = art.find_element(By.CSS_SELECTOR, 'div.sdp-review__article__list__review > div').text
                    content = re.sub(r"[\n\t]", "", content.strip())
                except NoSuchElementException:
                    content = ""

                try:
                    rating = int(art.find_element(By.CSS_SELECTOR, 'div.sdp-review__article__list__info__product-info__star-orange').get_attribute('data-rating'))
                except NoSuchElementException:
                    rating = 0  # 혹은 0

                reviews.append({
                        "작성자": username,
                        "작성일": date,
                        "평점": rating,
                        "리뷰내용": re.sub(r"[\n\t]", "", content.strip())
                    })
                print(f"{len(reviews):3}. 작성자: {username}")
                print(f" 현재 수집한 리뷰 수: {len(reviews)} / 현재 페이지: {current_page}")
            current_page = click_next_page(driver, current_page)
    except Exception as e:
        driver.quit()
        return reviews
    finally:
        try:
            time.sleep(0.1)
            driver.close()
        except:
            pass
        return reviews                             
                
                
if __name__ == "__main__":
    url = input("크롤링할 쿠팡 상품 URL을 입력하세요:\n").strip()
    driver = setup_driver()
    if "coupang.com" not in url:
        print("❌ 유효한 쿠팡 상품 URL이 아닙니다.")
        driver.quit()
    else:
        review_data = crawl_reviews(url,driver)
        if review_data:
            save_to_excel(review_data)
        driver.quit()





