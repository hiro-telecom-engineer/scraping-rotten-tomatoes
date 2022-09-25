import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import datetime
import pprint

URL = "https://www.rottentomatoes.com/m/halloween_kills/reviews?type=user"
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option(    "excludeSwitches", ["enable-automation"])
driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

review_dic = {"reviewer":None,"date":None,"evaluation":None,"review":None,"reviewer_url":None}
fail_counter = {"main":0,"get_inf":0,"get_inf_for":0}

def main(url):
    # 初期化
    excel = make_excel(review_dic)
    # 確認したい映画URLへアクセス
    driver.get(url)
    time.sleep(1)
    # 全レビューページの取得
    for i in range(1000000):
        get_inf(excel)
        try:
            element = driver.find_element_by_xpath("//*[@id='content']/div/div/nav[3]/button[2]")
            if "hide" not in element.get_attribute("class"):
                element.click()
                time.sleep(1)
            else:
                print("検索終了")
                break
        except NoSuchElementException:
            fail_counter["main"] += 1
            pass
    now = datetime.datetime.now()
    file_name = 'rotten_tomatoes_review_{}.xlsx'.format(now.strftime('%Y%m%d_%H%M%S'))
    excel.save_file(file_name,"review")
    print(fail_counter)
    return


# ページ内レビューの情報取得
def get_inf(excel):
    try:
        all_review = driver.find_element_by_css_selector(".audience-reviews")
        review_lists = all_review.find_elements_by_tag_name('li')
        for review_list in review_lists:
            try:
                # ユーザ名
                review_dic["reviewer"] = review_list.find_element_by_css_selector(".audience-reviews__name-wrap").text
                # ユーザURL
                if review_list.find_elements_by_tag_name('a'):
                    review_dic["reviewer_url"] = review_list.find_element_by_tag_name('a').get_attribute("href")
                else:
                    review_dic["reviewer_url"] = None
                # 日付
                review_dic["date"] = review_list.find_element_by_css_selector(".audience-reviews__duration").text
                # 評価
                point = 0
                point_lists = review_list.find_elements_by_tag_name('span')
                for point_list in point_lists:
                    if "star-display__filled " == point_list.get_attribute("class"):
                        point += 1
                    if "star-display__half " == point_list.get_attribute("class"):
                        point += 0.5
                review_dic["evaluation"] = point
                # レビュー
                review = review_list.find_element_by_css_selector(".audience-reviews__review.js-review-text.clamp.clamp-8.js-clamp").text
                review_dic["review"] = review.replace(". ",".\n")
                excel.add_inf(review_dic)
                pprint.pprint(review_dic)
            except NoSuchElementException:
                fail_counter["get_inf_for"] += 1
                pass
    except NoSuchElementException:
        fail_counter["get_inf"] += 1
        pass
    return


class make_excel:
    # エクセルヘッダ記入
    def __init__( self , init_dic) :
        key_list = init_dic.keys()
        self.df = pd.DataFrame(columns=key_list)
    # 行追加
    def add_inf ( self , add_dict ) :
        self.df = self.df.append( add_dict , ignore_index=True)
        return
    # ファイル保存
    def save_file( self , file_name , title ):
        self.df.to_excel(file_name, sheet_name=title)
        return

if __name__ == "__main__":
    main(URL)
    driver.quit()