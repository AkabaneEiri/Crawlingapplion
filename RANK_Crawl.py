import os
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By

import time


# 기존 크롤링 함수는 그대로 유지
def scroll_page(driver, max_scrolls=10):
    """페이지를 반복적으로 스크롤하여 데이터를 강제로 로드."""
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  # 스크롤 후 대기
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:  # 더 이상 로드할 데이터가 없으면 종료
            break
        last_height = new_height


def get_rankings(driver, url, max_rank=20, platform="android"):
    """특정 URL에서 게임 순위를 크롤링."""
    driver.get(url)
    time.sleep(5)

    # 스크롤로 강제로 데이터 로드
    scroll_page(driver, max_scrolls=10)

    games = []

    # 플랫폼에 따라 랭킹 색상 선택자 변경
    rank_selectors = {
        "android": ["span.icon_rank.android_color", "div.rank", "span.rank"],
        "ios": ["span.icon_rank.iphone_color", "div.rank", "span.rank"]
    }
    name_selectors = ["p.blog", "h2.title", "div.title"]

    selectors = ["li.item", "div.item", "ul.itemList li", "div.name.wrap", "div.list-item"]

    for selector in selectors:
        try:
            game_items = driver.find_elements(By.CSS_SELECTOR, selector)
            if game_items:
                print(f"Found items using selector: {selector}")
                break
        except:
            continue

    if game_items:
        for item in game_items[:max_rank]:
            try:
                # 순위
                rank = None
                for rank_selector in rank_selectors[platform]:
                    try:
                        rank_element = item.find_element(By.CSS_SELECTOR, rank_selector)
                        rank = rank_element.text.strip()
                        if rank:
                            break
                    except:
                        continue

                # 이름
                name = None
                for name_selector in name_selectors:
                    try:
                        name_element = item.find_element(By.CSS_SELECTOR, name_selector)
                        name = name_element.text.strip()
                        if name:
                            break
                    except:
                        continue

                # 데이터 추가
                if rank and name:
                    games.append({"순위": int(rank), "이름": name})
            except Exception as e:
                print(f"Error processing item: {e}")
                continue

    return games


def compare_with_previous(current_file, save_path):
    previous_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
    previous_file = os.path.join(save_path, f'game_rankings_{previous_date}.xlsx')

    if not os.path.exists(previous_file):
        print(f"전날 파일을 찾을 수 없습니다: {previous_file}")
        return

    current_data = pd.ExcelFile(current_file)
    previous_data = pd.ExcelFile(previous_file)

    comparison_results = {}

    for sheet in current_data.sheet_names:
        if sheet in previous_data.sheet_names:
            current_df = current_data.parse(sheet).set_index("순위")
            previous_df = previous_data.parse(sheet).set_index("순위")

            for category in ["무료 이름", "유료 이름", "매출 이름"]:
                if category not in current_df.columns or category not in previous_df.columns:
                    continue

                # 순위 변동 계산
                current_temp = current_df.reset_index()
                previous_temp = previous_df.reset_index()

                # 추가된 게임
                current_temp["변동폭"] = "추가"  # 기본값은 "추가"로 설정
                previous_names = previous_temp[category].tolist()
                current_temp.loc[current_temp[category].isin(previous_names), "변동폭"] = 0  # 공통된 항목 초기화

                # 변동폭 계산
                common_games = current_temp[current_temp[category].isin(previous_temp[category])]
                merged = common_games.merge(
                    previous_temp, on=category, suffixes=("_현재", "_전날")
                )
                merged["변동폭"] = merged["순위_전날"] - merged["순위_현재"]

                # 결과 저장
                current_temp.update(merged[["변동폭"]])
                current_temp = current_temp[["순위", category, "변동폭"]]

                if sheet not in comparison_results:
                    comparison_results[sheet] = {}
                comparison_results[sheet][category] = current_temp

    # 결과 저장
    result_filename = os.path.join(save_path, f'comparison_results_{datetime.now().strftime("%Y%m%d")}.xlsx')
    with pd.ExcelWriter(result_filename, engine="openpyxl") as writer:
        for country, categories in comparison_results.items():
            combined_df = pd.DataFrame()
            for category, df in categories.items():
                if not df.empty:
                    df = df.set_index("순위")  # 순위를 인덱스로 사용
                    combined_df = pd.concat([combined_df, df], axis=1)

            if not combined_df.empty:
                # 열 순서 지정
                column_order = [
                    "무료 이름", "변동폭",
                    "유료 이름", "변동폭",
                    "매출 이름", "변동폭"
                ]
                combined_df.columns = pd.MultiIndex.from_tuples([(col, "") if col == "변동폭" else (col, "이름") for col in combined_df.columns])
                combined_df.to_excel(writer, sheet_name=country, index=True)

    print(f"비교 결과가 저장되었습니다: {result_filename}")


def crawl_game_rankings():
    """국가별 순위를 크롤링하고 엑셀에 저장."""
    # 저장 경로 설정
    home_dir = os.path.expanduser("~")
    save_path = os.path.join(home_dir, "Desktop", "gamerank")

    if not os.path.exists(save_path):
        os.makedirs(save_path)

    # 브라우저 옵션 설정
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')

    driver = webdriver.Chrome(options=options)

    try:
        # 국가별 URL 설정
        urls = {
             "한국 안드로이드": {
                "무료": ["https://applion.jp/android/rank/kr/6014/", "https://applion.jp/android/rank/kr/6014/?start=20"],
                "유료": ["https://applion.jp/android/rank/kr/6014/paid/", "https://applion.jp/android/rank/kr/6014/paid/?start=20"],
                "매출": ["https://applion.jp/android/rank/kr/6014/gross/", "https://applion.jp/android/rank/kr/6014/gross/?start=20"]
            },
            "일본 iOS": {
                "무료": ["https://applion.jp/iphone/rank/jp/6014/", "https://applion.jp/iphone/rank/jp/6014/?start=20"],
                "유료": ["https://applion.jp/iphone/rank/jp/6014/paid/", "https://applion.jp/iphone/rank/jp/6014/paid/?start=20"],
                "매출": ["https://applion.jp/iphone/rank/jp/6014/gross/", "https://applion.jp/iphone/rank/jp/6014/gross/?start=20"]
            },
            "미국 안드로이드": {
                "무료": ["https://applion.jp/android/rank/us/6014/", "https://applion.jp/android/rank/us/6014/?start=20"],
                "유료": ["https://applion.jp/android/rank/us/6014/paid/", "https://applion.jp/android/rank/us/6014/paid/?start=20"],
                "매출": ["https://applion.jp/android/rank/us/6014/gross/", "https://applion.jp/android/rank/us/6014/gross/?start=20"]
            },
            "미국 iOS": {
                "무료": ["https://applion.jp/iphone/rank/us/6014/", "https://applion.jp/iphone/rank/us/6014/?start=20"],
                "유료": ["https://applion.jp/iphone/rank/us/6014/paid/", "https://applion.jp/iphone/rank/us/6014/paid/?start=20"],
                "매출": ["https://applion.jp/iphone/rank/us/6014/gross/", "https://applion.jp/iphone/rank/us/6014/gross/?start=20"]
            }
        }

        # 엑셀 파일 생성
        current_date = datetime.now().strftime('%Y%m%d')
        filename = os.path.join(save_path, f'game_rankings_{current_date}.xlsx')

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for country, categories in urls.items():
                print(f"\n크롤링 시작: {country}")
                country_data = {}

                # 플랫폼 구분
                platform = "ios" if "iOS" in country else "android"

                for category, url_list in categories.items():
                    combined_data = []
                    for url in url_list:
                        data = get_rankings(driver, url, max_rank=20, platform=platform)
                        combined_data.extend(data)

                    if combined_data:
                        df = pd.DataFrame(combined_data)
                        df['순위'] = pd.to_numeric(df['순위'])
                        df = df.sort_values('순위').set_index('순위')
                        country_data[category] = df['이름']

                # 국가별 데이터 병합
                if country_data:
                    merged_df = pd.concat(country_data, axis=1)
                    merged_df.columns = [f"{col} 이름" for col in merged_df.columns]
                    merged_df.to_excel(writer, sheet_name=country, index=True)
                    print(f"{country} 데이터 저장 완료")

        print(f"\n엑셀 파일이 저장되었습니다: {filename}")
        return filename

    finally:
        driver.quit()


# 크롤링 실행 및 비교
filename = crawl_game_rankings()
if filename:
    print(f"데이터가 저장되었습니다: {filename}")
    home_dir = os.path.expanduser("~")
    save_path = os.path.join(home_dir, "Desktop", "gamerank")
    compare_with_previous(filename, save_path)
else:
    print("크롤링에 실패했습니다.")
