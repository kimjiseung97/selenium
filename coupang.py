from datetime import datetime
import math
import time
import re
from tkinter import ttk
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, NoSuchElementException
from openpyxl import Workbook
from selenium.webdriver.chrome.service import Service as ChromeService
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading

class ReviewApp:
    def __init__(self, root):
        self.root = root
        self.root.title("쿠팡 리뷰 수집기")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        # 입력 요소를 담을 프레임
        form_frame = tk.Frame(root)
        form_frame.pack(pady=20)

        # URL 입력
        tk.Label(form_frame, text="쿠팡 상품 URL:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.url_entry = tk.Entry(form_frame, width=40)
        self.url_entry.grid(row=0, column=1, pady=5)

        # 정렬 기준
        tk.Label(form_frame, text="리뷰 정렬 기준:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.sort_option = ttk.Combobox(form_frame, values=["베스트순", "최신순"], state="readonly", width=37)
        self.sort_option.current(0)
        self.sort_option.grid(row=1, column=1, pady=5)

        # 수집 리뷰 갯수
        tk.Label(form_frame, text="수집 리뷰 갯수:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.review_count = ttk.Combobox(form_frame, values=["100", "200", "300", "400", "500", "600", "700","800","900","1000","1100","1200","1300","1400","1500"], state="readonly", width=37)
        self.review_count.current(0)
        self.review_count.grid(row=2, column=1, pady=5)

        # 실행 버튼 예시
        self.start_btn = tk.Button(root, text="리뷰 수집 시작")
        # self.start_btn.pack(pady=10)
        # self.review_count = ttk.Combobox(
        #     root,
        #     values=[str(i) for i in range(100, 1600, 100)],
        #     state="readonly"
        # )
        # self.review_count.current(0)
        # self.review_count.pack(pady=5)
        # self.url_entry = tk.Entry(root, width=80)
        # self.url_entry.pack(pady=5)

        self.start_button = tk.Button(root, text="리뷰 수집 시작", command=self.start_scraping)
        self.start_button.pack(pady=5)

        self.log_area = scrolledtext.ScrolledText(root, width=80, height=15)
        self.log_area.pack(padx=10, pady=10)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)        

    def start_scraping(self):
        url = self.url_entry.get().strip()
        sort = self.sort_option.get()
        count = self.review_count.get()  
        

        if "coupang.com" not in url:
            messagebox.showerror("URL 오류", "쿠팡 상품 URL을 입력하세요.")
            return
        
        
        
        count = int(count)        
        thread = threading.Thread(target=self.scrape_reviews, args=(url,sort,count))
        thread.start()

    def scrape_reviews(self, url, sort,count):
        self.log("크롬 드라이버 시작 중...")
        driver = setup_driver()
        self.log("리뷰 수집 시작...")

        try:
            reviews ,product_name= crawl_reviews(url, driver, log_func=self.log,sort=sort,count=count)
            if reviews and product_name:
                self.log(f"총 수집된 리뷰 수: {len(reviews)}")
                save_to_excel(reviews,product_name)
                self.log("엑셀 저장 완료: " + product_name)
            else:
                self.log("리뷰를 수집하지 못했습니다.")
        except Exception as e:
            self.log(f"에러 발생: {e}")


def setup_driver():
    options = uc.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--start-maximized')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    driver = uc.Chrome(options=options)
    return driver

def save_to_excel(reviews,priduct_name):
    wb = Workbook()
    ws = wb.active
    today_str = datetime.today().strftime("%Y%m%d")
    ws.append(["작성자", "작성일", "평점", "리뷰 내용"])
    for r in reviews:
        ws.append([r["작성자"], r["작성일"], r["평점"], r["리뷰내용"]])
    wb.save(priduct_name + "_" +today_str + ".xlsx")
    
    
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



        

def crawl_reviews(url,driver,log_func=print,sort="베스트순",count=100):
    driver.get(url)
    time.sleep(2)
    
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "상품평")]'))
        ).click()
        time.sleep(2)

        sort_xpath_map = {
            "베스트순": "//div[@class='review-order-container']/button[normalize-space(text())='베스트순']",
            "최신순": "//div[@class='review-order-container']/button[normalize-space(text())='최신순']"
        }
        
        sort_xpath = sort_xpath_map.get(sort)
        if sort_xpath:
            log_func(f"정렬 기준 선택: {sort}")
            sort_tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, sort_xpath))
            )
            driver.execute_script("arguments[0].click();", sort_tab)
            time.sleep(1.5)
    except UnexpectedAlertPresentException:
        try:
            alert = driver.switch_to.alert
            log_func(f"쿠팡에서 차단되었습니다. Alert 메시지: {alert.text}")
            alert.accept()
        except:
            pass
        driver.quit()
        return []
    except Exception as e:
        log_func(f"리뷰 탭 진입 실패: {e}")
        driver.quit()
        return []

    reviews = []
    
    try:
        product_name = driver.find_element(By.XPATH, "//h1[contains(@class, 'product-title')]/span").text.strip()
    except Exception:
        product_name = "상품명 없음"
        
    try:
        
        total_review_count = get_review_totalcount(driver)
        if count :
            max_review_count = count
            
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
                log_func(f"현재 수집한 리뷰 수: {len(reviews)} / 현재 페이지: {current_page}")
                if len(reviews) >= max_review_count:
                    log_func("최대 리뷰 개수에 도달했습니다 크롤링을 종료합니다")
                    return reviews , product_name
            current_page = click_next_page(driver, current_page)
    except Exception as e:
        driver.quit()
        return reviews , product_name
    finally:
        try:
            time.sleep(0.1)
            driver.close()
        except:
            pass
        return reviews,product_name                            
                
                
if __name__ == "__main__":
    root = tk.Tk()
    app = ReviewApp(root)
    root.mainloop()





