from optparse import Option
import time
import os

from numpy import mat
from const import all_const as cnst
from common import common_method as cmn

from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

from PIL import ImageGrab

# 参考 - requestsライブラリを利用して、ダウンロードURLとCockieでファイルを取得）
# https://news.mynavi.jp/techplus/article/zeropython-50/

class crawler_org :

    # ハードコピー用
    sheet_name = ''

     # 画面ショットを撮る
    def take_display(self, driver, image_name) :
        cmn.write_log('takeDisplay() 開始')

        # get width and height of the page
        w = driver.execute_script("return document.body.scrollWidth;")
        h = driver.execute_script("return document.body.scrollHeight;")

        # set window size
        driver.set_window_size(w,h)

        # ディレクトリ確認
        if not os.path.exists(cnst.IMAGE_DIR):
            os.mkdir(cnst.IMAGE_DIR)
            cmn.write_log('ディレクトリ[' + cnst.IMAGE_DIR + ']を作成しました')

        # Get Screen Shot
        driver.save_screenshot(cnst.IMAGE_DIR + image_name + ".png")
        cmn.write_log('takeDisplay() 終了')



    # 画面ショットを撮る（画面全体、JavaScript用）
    def take_display_all(self, image_name) :
        cmn.write_log('takeDisplayAll() 開始')
        img = ImageGrab.grab()
        img.save(cnst.IMAGE_DIR + image_name + ".png")
        cmn.write_log('takeDisplayAll() 終了')



    # 画像のファイル名を生成する
    def get_image_file_name(self, current_sheet_name:str, img_no:str )-> str:
        return current_sheet_name + '_' + str(img_no).zfill(3) 



    # Chrome Driver 設定
    def get_browser_option(self) -> Options:
        opt = Options()
        opt.add_experimental_option('prefs', { "download.prompt_for_download": False 
                                                ,"download.directory_upgrade": False
                                                ,"download.manager.showWhenStarting": False
                                                ,"download.default_directory": cnst.TEMP_DIR })
        
        # opt.add_argument('--headless')          # ブラウザを表示しない
        opt.add_argument('--hide-scrollbars')   # スクロールバーを非表示にする
        opt.add_argument('--incognito')         # シークレットモードでChromeを起動する

        return opt



    # ブラウザ操作
    def operate_control(self, driver:webdriver.Chrome ,row:dict):
        cmn.write_log('operate_control() 開始' )
        try :
            # エレメント（オブジェクト）を取得
            print(100)
            match row['particular'] :
                case 'ID値':
                    print(110)
                    element = driver.find_element_by_id(row['object_name'])
                case 'Name値':
                    print(120)
                    element = driver.find_element_by_Name(row['object_name'])
                case 'Xpath':
                    print(130)
                    element = driver.find_element_by_xpath(row['object_name'])
                case _ :
                    print(140)

            print(200)
            if element == None : 
                cmn.write_log('エレメント取得エラー（値:['+row['particular']+']）')
                return

            # エレメントの操作
            print(300)
            match row['handling'] :
                case '入力' :
                    print(310)
                    element.clear()
                    element.send_keys(str(row['value']))

                case '選択(順番)' :
                    print(320)
                    select = Select(element)
                    select.select_by_index(str(row['value']))

                case '選択(Value値)' :
                    print(330)
                    select = Select(element)
                    select.select_by_value(str(row['value']))

                case '選択(表示値)' :
                    print(340)
                    select = Select(element)
                    select.select_by_visible_text(str(row['value']))

                case 'クリック' :
                    print(350)
                    element.click()

                case _ :
                    print(390)
                    cmn.write_log('ハンドリング取得エラー（値:['+row['handling']+']）')
                    return
            print(900)
        except Exception as e :
            raise e 
        cmn.write_log('operate_control() 終了' )



    # ---- メイン操作 ------
    def run_brawser(self, df:list) :
        cmn.write_log('run_brawser() 開始')
        try:
            # Chrome Driverの設定を行う
            print(10)
            opt = self.get_browser_option()
            driver = webdriver.Chrome(executable_path=cnst.CHROME_DRIVER,
                                        chrome_options=opt)
            for row in df:
                print(20)
                cmn.write_log('シート=[' + row['sheet_name'] + ']  行番目=['+ str(row['row_number']) + ']')
                # rowはdict型
                match row['control']:
                    case 'URL' :
                        print(21)
                        driver.get(row['value'])

                    case 'jsダイアログ(OKボタン)':
                        print(22)
                        Alert(driver).accept()

                    case 'jsダイアログ(Cancelボタン)':
                        print(23)
                        Alert(driver).dismiss()

                    case 'ハードコピー' :
                        print(24)
                        self.take_display(driver, self.get_image_file_name(row['sheet_name'], row['row_number'])) 

                    case 'ハードコピー(全体)' :
                        print(25)
                        self.take_display_all(self.get_image_file_name(row['sheet_name'], row['row_number'])) 

                    case '画面切り替え' :
                        print(26)
                        handle_array = driver.window_handles
                        max_len = len(handle_array)
                        driver.switch_to.window(handle_array[max_len - 1])

                    case '待機(秒)' :
                        print(27)
                        time.sleep(int(row['value']))

                    case _:
                        print(28)
                        # アクセスする
                        self.operate_control(driver, row)

                time.sleep(2)
                print(90)
       
            driver.quit()

        except Exception as e:
            cmn.write_log("run_brawser() Exception検知")
            raise e
        cmn.write_log('run_brawser() 終了')

