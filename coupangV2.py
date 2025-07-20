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
from contextlib import contextmanager
import atexit

# 전역 드라이버 리스트 (종료 시 정리용)
_active_drivers = []

def cleanup_drivers():
    """프로그램 종료 시 모든 드라이버 정리"""
    global _active_drivers
    for driver in _active_drivers:
        try:
            if driver:
                safe_quit_driver(driver)
        except:
            pass
    _active_drivers.clear()

# 프로그램 종료 시 자동 정리 등록
atexit.register(cleanup_drivers)

def safe_quit_driver(driver):
    """드라이버를 안전하게 종료"""
    if not driver:
        return
    
    try:
        # 모든 창 닫기
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            driver.close()
    except:
        pass
    
    try:
        driver.quit()
    except:
        pass
    
    # 서비스 프로세스 정리
    try:
        if hasattr(driver, 'service') and driver.service:
            driver.service.stop()
    except:
        pass

@contextmanager
def setup_driver():
    """컨텍스트 매니저로 드라이버 관리"""
    driver = None
    global _active_drivers
    
    try:
        options = uc.ChromeOptions()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--disable-logging')
        options.add_argument('--disable-gpu-logging')
        options.add_argument('--log-level=3')
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
        
        driver = uc.Chrome(options=options)
        _active_drivers.append(driver)
        print("✅ Chrome 드라이버가 성공적으로 시작되었습니다.")
        
        yield driver
        
    except Exception as e:
        print(f"❌ 드라이버 생성 중 오류 발생: {e}")
        yield None
    finally:
        if driver:
            try:
                # 활성 드라이버 목록에서 제거
                if driver in _active_drivers:
                    _active_drivers.remove(driver)
                safe_quit_driver(driver)
                print("✅ Chrome 드라이버가 안전하게 종료되었습니다.")
            except Exception as e:
                print(f"⚠️ 드라이버 종료 중 오류 (무시 가능): {e}")

def save_to_excel(reviews, filename="coupang_reviews.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.append(["작성자", "작성일", "평점", "리뷰 내용"])
    for r in reviews:
        ws.append([r["작성자"], r["작성일"], r["평점"], r["리뷰내용"]])
    wb.save(filename)
    print(f"✅ 리뷰 데이터가 {filename}에 저장되었습니다.")
    
def get_review_totalcount(driver):        
    try:
        review_count_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[contains(text(), "상품평")]'))
        )
        text = review_count_element.text
        match = re.search(r'([\d,]+)', text)
        total_reviews = int(match.group(1).replace(',', '')) if match else 0
        return total_reviews
    except:
        return 0
    
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
        driver.execute_script("arguments[0].click();", next_page_btn)
        current_page += 1
        time.sleep(1.5)

        # 페이지 변경 확인
        WebDriverWait(driver, 5).until(
            lambda d: any(
                e.text not in old_usernames 
                for e in d.find_elements(By.CSS_SELECTOR, 'span.sdp-review__article__list__info__user__name')
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
                    e.text not in old_usernames 
                    for e in d.find_elements(By.CSS_SELECTOR, 'span.sdp-review__article__list__info__user__name')
                )
            )
        except:
            current_page += 1  # 실패해도 무한루프 방지용 강제 증가
    
    return current_page

def crawl_reviews(url, driver):
    
    reviews = []
    
    try:
        print(f"🔄 URL 접속 중: {url}")
        options = uc.ChromeOptions()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--disable-logging')
        options.add_argument('--disable-gpu-logging')
        options.add_argument('--log-level=3')
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
        driver = uc.Chrome(options=options)
    
        driver.get(url)
        time.sleep(3)
        
        try:
            WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "상품평")]'))
            ).click()
            time.sleep(3)
            print("✅ 상품평 탭 클릭 완료")
        except UnexpectedAlertPresentException:
            try:
                alert = driver.switch_to.alert
                print("❌ 쿠팡에서 차단되었습니다. Alert 메시지:", alert.text)
                alert.accept()
            except:
                pass
            return []
        except Exception as e:
            print(f"❌ 리뷰 탭 진입 실패: {e}")
            return []

        try:
            total_review_count = get_review_totalcount(driver)
            total_pages = math.ceil(total_review_count / 10)
            current_page = 1
            
            print(f"📊 총 리뷰 수: {total_review_count}, 총 페이지 수: {total_pages}")
            
            # 최대 50페이지로 제한 (너무 많은 페이지 방지)
            max_pages = min(total_pages, 50)
            
            for page in range(1, max_pages + 1):
                try:
                    articles = WebDriverWait(driver, 15).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'article.sdp-review__article__list'))
                    )
                    print(f"📄 페이지 {page}에서 {len(articles)}개 리뷰 발견")
                except:
                    print(f"⚠️ 페이지 {page}에서 리뷰를 찾을 수 없습니다.")
                    continue

                page_reviews = 0
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
                    except (NoSuchElementException, ValueError, TypeError):
                        rating = 0

                    reviews.append({
                        "작성자": username,
                        "작성일": date,
                        "평점": rating,
                        "리뷰내용": re.sub(r"[\n\t]", "", content.strip())
                    })
                    page_reviews += 1
                
                print(f"📄 페이지 {current_page} 완료 - 이번 페이지: {page_reviews}개, 총 수집: {len(reviews)}개")
                
                # 마지막 페이지가 아니면 다음 페이지로 이동
                if page < max_pages:
                    try:
                        current_page = click_next_page(driver, current_page)
                        time.sleep(2)  # 페이지 이동 후 대기
                    except Exception as e:
                        print(f"⚠️ 다음 페이지 이동 실패: {e}")
                        break
                        
        except Exception as e:
            print(f"❌ 리뷰 수집 중 오류 발생: {e}")
            
    except Exception as e:
        print(f"❌ 전체 크롤링 과정에서 오류 발생: {e}")
    
    return reviews

def main():
    """메인 실행 함수"""
    try:
        url = input("크롤링할 쿠팡 상품 URL을 입력하세요:\n").strip()
        
        if "coupang.com" not in url:
            print("❌ 유효한 쿠팡 상품 URL이 아닙니다.")
            return
        
        with setup_driver() as driver:
            if driver:
                review_data = crawl_reviews(url, driver)
                
                if review_data:
                    print(f"🎉 총 {len(review_data)}개의 리뷰를 수집했습니다!")
                    save_to_excel(review_data)
                else:
                    print("❌ 수집된 리뷰가 없습니다.")
            else:
                print("❌ 드라이버 초기화에 실패했습니다.")
                
    except KeyboardInterrupt:
        print("\n⚠️ 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"❌ 프로그램 실행 중 오류: {e}")
    finally:
        cleanup_drivers()
        print("🔄 모든 리소스가 정리되었습니다.")

if __name__ == "__main__":
    main()