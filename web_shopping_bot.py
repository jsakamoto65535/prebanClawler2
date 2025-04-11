from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
import time
from selenium.webdriver.common.keys import Keys
import datetime
import schedule
import re

# 環境変数の読み込み
load_dotenv()

class WebShoppingBot:
    def __init__(self):
        # 設定ファイルから取得
        self.targetList = os.getenv('TARGET_LIST').split(',')
        self.preCheckList = os.getenv('PRECHECK_LIST').split(',')
        
        self.sender_email = os.getenv('EMAIL_USER')
        self.sender_password = os.getenv('EMAIL_PASSWORD')
        self.receiver_email = os.getenv('RECEIVER_EMAIL')
        self.retryMax = int(os.getenv('ADD_CART_RETRY_COUNT'))
        self.isHeadLessMode = bool(int(os.getenv('BACKGOUND_MODE')))

        # Chromeオプションの設定
        chrome_options = Options()
        # ヘッドレスモードを有効化選択
        if self.isHeadLessMode:
            chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36')
        
        #ブラウザが自動化されていることを
        chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        # WebDriverの初期化
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        
        # 初期値設定
        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587
        self.addCartCount = 0
        self.isAddedCart = False
        
    #カートに追加処理
    def checkItemNmeAndAddToCart(self, foundItemElement):
        foundItemName = foundItemElement.text
        foundItemName = foundItemName.replace('\n', '')

        # これはURLの文字列が返るみたい
        linkItem = foundItemElement.get_attribute("href")
        # URLを開くにはclickではなくgetを使う
        self.driver.get(linkItem)
        time.sleep(1)
        #見つかったアイテムをカートに入れる
        addCartRetryCount = 0
        # 戻り値用カート追加フラグ
        isAddCart = False
        
        # 詳細商品名チェック:商品名を取得
        detailItemName = self.driver.find_element(By.XPATH, f'//*[@id="box"]')
        foundInTargetList=False
        # 商品ページの詳細商品名と一致するものがなければスキップ
        for targetName in self.targetList: 
            if targetName in detailItemName.text:
                foundInTargetList =True
                break
        if foundItemName == False:        
            return isAddCart
        
        while addCartRetryCount < self.retryMax:
            try:
                # 予約ボタンの要素を取得
                buyElement = self.driver.find_element(By.XPATH, f'//*[@id="buy_side"]')
                #在庫がない場合はリロードする。クリックは失敗しないので文字列チェック
                if buyElement.text == '在庫がありません':
                    raise ValueError(f'在庫なし:{foundItemName}')
                
                #buyElement.send_keys(Keys.ENTER)
                buyElement.click()
                time.sleep(2)
                # カートに一個でも追加できた
                isAddCart = True
                self.isAddedCart = True
                self.addCartCount+=1
                
                # 追加した商品を通知に追加
                self.mail_body += '購入成功:' + foundItemName + '\n'
                break                        
            except Exception as buyException:
                print('在庫がない又はその他の例外')
                #物がなければリロード
                self.driver.refresh()
                time.sleep(2)
            addCartRetryCount+=1
            #リトライ時は少し待つ
            time.sleep(1)
                        
        if( addCartRetryCount >= self.retryMax) :
            self.mail_body += '購入失敗：' + foundItemName + '\n'
            print(f'cant buy {linkItem.title}')
        # 購入出来たら前のページに戻る
        self.driver.back()
        time.sleep(1)
        return isAddCart
        
 

    #
    #   新着情報ページで商品を検索してカートに入れる
    #
    def crawler_items(self):
        """新着情報ページの要素を検索してカートに入れる"""
        try:
            #開始前にリロード
            self.driver.refresh()
            time.sleep(2)
            listCount = 1
            # 新着商品1ブロック目の配列だけチェックする。要素数はわからないのでmax100個見る。なければ例外で次に行く
            while listCount < 100:
                # 1ブロック目の配列を順次見ていく
                allNewLink = self.driver.find_element(By.XPATH, f'//*[@id="cdu2mainColumn"]/div/ul[1]/li[{listCount}]/p/a')                
                checkItem = allNewLink.text.replace('\n','')
                # 事前チェックリスト分チェックする。このページでは商品名が省略表示されるため事前にフィルタをかける
                for targetWord in self.preCheckList:
                    # ターゲットリストにある商品名であればカートに入れる
                    if targetWord in allNewLink.text:
                        print(f"事前チェックで検出[{listCount}]:{targetWord} -> {checkItem}")
                        # 商品ページを開き詳細な商品名をチェックし購入商品であればカートに入れる
                        if self.checkItemNmeAndAddToCart(allNewLink):
                            # カートに追加できたら、その商品をターゲットリストから削除
                            self.targetList.remove(targetWord)
                            break
                        
                print(f"NotFindItem[{listCount}]: {checkItem}")
                listCount += 1        
        #1個目の要素ブロックの最後までいったら例外で先に進む。正常処理           
        except Exception as e:
            #print(f"End of Items: {e}")
            print(f"add cart:{self.addCartCount}")

    #
    #   メール送信処理
    #
    def send_notification_email(self, product_name):
        """カートへの追加をメールで通知"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = self.receiver_email
            msg['Subject'] = product_name

            body = self.mail_body
            msg.attach(MIMEText(body, 'plain'))

            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)

            print("通知メールを送信しました")
        except Exception as e:
            print(f"メール送信中にエラーが発生しました: {e}")
    #
    #   処理終了
    #
    def close(self):
        """ブラウザを閉じる"""
        # self.driver.quit()

    #
    #   メイン処理：
    #
    def mainProcess(self):
        # 1回分の初期化
        self.isAddedCart = False
        self.mail_body ="プレバン通知\n"
        self.mail_body += "購入リスト：" + os.getenv('TARGET_LIST') +"\n"
        self.addCartCount = 0
        
        print('新着ページを表示して購入アイテムがあればカートに追加')
        # 新着情報ページで商品を検索しカートに追加
        self.crawler_items()
        
        # 結果をメール通知(メアドを設定していなければ送らない)
        if len(self.sender_email) != 0:
            subject = "プレバン通知:商品が見つかりませんでした"
            if self.isAddedCart == True:
                # メールタイトル
                subject = "プレバン通知:商品をカートに入れました"
            self.send_notification_email(subject)
    
    def test(self):
        print('exec')           



def main():
    bot = WebShoppingBot()
    try:
        # 環境変数から設定を読み込む
        base_url = os.getenv('SHOP_URL')
        clawle_times = os.getenv('EXECUTE_TIME').split(',')
        pollingTime = int(os.getenv('EXECUTE_SCHEDULE_POLLING_PERIOD'))
        scheduleMode = bool(int(os.getenv('SCHEDULE_MODE')))
        pollingTimeAtNormalMode = int(os.getenv('RETRY_WAIT'))
        # プレバンを表示
        # パスワードを入れたりし終わったらキーを押して次に進むようにする
        bot.driver.get(base_url)
        # 新着情報へのリンクをクリック
        new_arrivals_link = WebDriverWait(bot.driver, 10).until(
        #    EC.presence_of_element_located((By.LINK_TEXT, "すべて見る"))
        EC.presence_of_element_located                
        )
        #time.sleep(2)  # ページ読み込み待機
        #time.sleep(3*60)
        inputKey = input('ログインしてからこのページに戻ってきてください。そしてキーを押して先に進みます\n')

        print('add these item to cart')
        print(os.getenv('TARGET_LIST'))

        # スケジュールモード
        if scheduleMode == True:
            #実行時間を登録
            print(f'\nRegister Clawle Time{clawle_times}')
            for clawleTime in clawle_times:
                schedule.every().day.at(clawleTime).do(bot.mainProcess)
            
            #登録された時間で処理を起動する
            print('\nStart Waiting!!')
            while True:
                schedule.run_pending()
                time.sleep(pollingTime)        
        #即時モード：購入商品がなくなるまで繰り返す
        else:
            print('\nStart imediate mode!!')
            while True:
                print('メイン処理を開始')
                # スケジュールモードでない場合のシングル実行
                bot.mainProcess()
                #リストがなくなっていれば終了
                if len(bot.preCheckList) == 0:
                    print('全ての商品をカートに入れました。プログラムを終了します。')
                    break
                else:
                    print(f'メイン処理終了 {pollingTimeAtNormalMode}秒後リトライします')
                    time.sleep(pollingTimeAtNormalMode)

    #エラー表示
    except Exception as e:
        print(e)
        
    finally:
        bot.close()

if __name__ == "__main__":
    main() 