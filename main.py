import os
import requests # コードの実行にはrequestsパッケージのインストールが必要
import time
from datetime import datetime, timedelta
from datetime import datetime, timezone
import slackweb # コードの実行にはslackwebパッケージのインストールが必要
from selenium import webdriver # コードの実行にはseleniumパッケージのインストールが必要
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager # webdriver-managerのインストールが必要
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

#ノートPCかデスクトップか。ノートPCなら1デスクトップなら0
laptop = 1
# LINE Notifyのトークン
LINE_NOTIFY_TOKEN = "" # LINEnotifyで所得したトークン(グループに送信するためのトークン)を入力
LINE_NOTIFY_TOKEN_MINE = "" # 学生個別連絡を自分に送るためのトークン
if laptop == 0:
    LINE_NOTIFY_TOKEN = LINE_NOTIFY_TOKEN_MINE # テスト用
# slack通知のいろいろ
slack_member_ID = "" # Slackメンション用のID
notification_name = "Porta" # Slackで表示される通知のユーザーネーム
slack_url = "" # SlackWebHookで取得したそのチャンネルのURL
# ログイン情報
mailaddress = "" # 学籍番号（mとイニシャルも含める　例:m12345a）
password = "" # パスワード
# 通知検索
keywords = ["M2","全医学部生","全学","個別"] # 対象のキーワード（他の人も含む）
# 実行時刻記録用ファイルのパス
if laptop == 0 :
    time_record_file = r"donetime.txt" # donetime.txtは存在していなくても作成される。ディレクトリ（フォルダ）さえ正しく指定できれば問題ない。
else :
    time_record_file = r"donetime.txt"

# 健康推進センターのお知らせテキスト
healthcenter_text = r"健康推進センター.txt"

# 実行のONOFF切り替え用フォルダのパス
remote_switch_path = r"\Portaのスイッチ"

# 最後に実行した時刻を出力する関数
def check_starttime():
    # 現在時刻を取得
    global now 
    now = datetime.now()
    try :
        # txtファイルに記録されている時刻と現在時刻の差を出す
        with open(time_record_file,mode="r",encoding="utf-8") as t :
            s = t.read()
            print (f"前回の実行時刻：{s}")
            check_start_time = datetime.strptime(s,"%Y/%m/%d %H:%M:%S")
    except :
        check_start_time = now - timedelta(hours=1) # 前回の実行時刻が取得できない場合、現在時刻から1時間前までを確認する
    
    finally :
        # txtファイルに現在時刻を書き込む準備
        global donetime
        donetime = now.strftime("%Y/%m/%d %H:%M:%S")
        return check_start_time
    
# slackに通知 errorが1ならメンション
def slack_notify(message,error=0):
    slack = slackweb.Slack(url=slack_url)
    if error == 1:
        slack.notify(text=f"<@{slack_member_ID}>{message}",username = notification_name)
        print ("エラーをSlackに送信しました。")
    else:
        slack.notify(text=f"{message}",username = notification_name)

# LINE Notifyに画像を送信する関数
def LINE_Notify(token,message,image_path=None):
    url = "https://notify-api.line.me/api/notify"
    headers = {"Authorization": f"Bearer {token}"}
    data = {"message": message}
    if image_path:
        files = {"imageFile": open(image_path, "rb")}
        res = requests.post(url, headers=headers, data=data, files=files)
    else:
        res = requests.post(url, headers=headers, data=data)
    # 送信結果を表示
    print(res.status_code, res.json()['message'])
    ratelimit = res.headers.get("X-RateLimit-Limit") # 1時間に可能なAPI callの上限回数を取得
    ratelimit_remain = res.headers.get("X-RateLimit-Remaining API") # callが可能な残りの回数
    ratelimit_imagelimit = res.headers.get("X-RateLimit-ImageLimit") #1時間に可能なImage uploadの上限回数
    ratelimit_imageremain = res.headers.get("X-RateLimit-ImageRemaining") #Image uploadが可能な残りの回数
    ratelimit_reset = res.headers.get("X-RateLimit-Reset") #リセットされる時刻
    # UNIXエポック秒を datetime オブジェクトに変換
    reset_time = datetime.fromtimestamp(int(ratelimit_reset), tz=timezone.utc)
    # 日本時間（JST）に変換
    jst_time = reset_time.astimezone(timezone(timedelta(hours=9)))
    print(f"1時間に可能なAPI callの上限回数: {ratelimit}")
    print(f"callが可能な残りの回数: {ratelimit_remain}")
    print(f"1時間に可能なImage uploadの上限回数: {ratelimit_imagelimit}")
    print(f"Image uploadが可能な残りの回数: {ratelimit_imageremain}")
    print(f"リセット時刻（日本時間）: {jst_time.strftime('%Y-%m-%d %H:%M:%S %Z')}")

# ポータル開いてお知らせ検索して、開いてスクリーンショットしてLINEnotifyに送信する関数
def take_screenshot_and_send():
    # Chromeのヘッドレスモードを設定
    chrome_options = Options()
    chrome_options.add_argument("--headless")

    # ウィンドウサイズを指定（例：1920x1080）
    #chrome_options.add_argument("--window-size=1920,1080")
    
    # Webドライバーを初期化
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)

    k = [] #すでに撮影したもののインデックスを保管するリスト
    notices_count =0 #通知の数を記録しておく変数
    j = 0 #すべて撮影終了したかどうかのステータス
    shot = 0 #スクリーンショットを撮影したかどうかのステータス
    error = 0 #エラーが生じたかどうかのステータス
    
    try:
        try:
            # 大学のポータルサイトにアクセス
            driver.get("https://ep.med.toho-u.ac.jp/")
            print("ページにアクセスしました")
        
            # ページが完全に読み込まれるのを待つ
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "MAILADDRESS"))
            )
            print("ログインページが読み込まれました")
        except:
            slack_notify("ポータルサイトにアクセスできなかった…",1)
            error = 1
            return
        
        try:
            # ログイン処理
            username_field = driver.find_element(By.NAME, "MAILADDRESS")
            password_field = driver.find_element(By.NAME, "LOGINPASS")
            login_button = driver.find_element(By.XPATH, "//input[@type='submit' and @value='ログイン']")
            username_field.send_keys(mailaddress)
            password_field.send_keys(password)
            time.sleep(1)
            login_button.click()
            print("ログイン情報を入力し、ログインボタンをクリックしました")
            
            # ページが読み込まれるのを待つ
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            print("ログイン後のページが読み込まれました")
        except:
            slack_notify("ログインできなかった…",1)
            error = 1
            return
        
        # テーブルが存在するまで待機
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//table[@id='T1']"))
            )
            print("テーブルが見つかりました")
        except TimeoutException:
            print("テーブルが見つかりませんでした")
            slack_notify("テーブルが見つからなかった…",1)
            error = 1
            return

        check_start_time = check_starttime() # 通知対象に含める期間
        current_year = now.year #今年
        
        Table_list = ["T1","T2","T3","T4"] # お知らせの検索をかけるテーブルのidリスト
        for table_name in Table_list :
            print(f"Table{table_name}を確認します。")

            while j == 0 :
                notices = driver.find_elements(By.XPATH, f"//table[@id='{table_name}']//tr")
                del notices[:4] # お知らせテーブルの最初のヘッダーなどをリストから削除
                print(f"見つかった通知の数: {len(notices)}")

                # 通知がない場合そのテーブルをパスする
                if len(notices) == 0 :
                    print(f"{table_name}には通知がありませんでした。")
                    break

                # 通知の数が変わった場合、撮影済みリストをその分ずらす
                if len(notices) != notices_count :
                    for i,l in enumerate(k) :
                        k[i] = l + len(notices) - notices_count
                notices_count = len(notices)
                print (f"撮影済みリストの中身は{k}")

                for i, notice in enumerate(notices,start=1):
                    try:
                        # tdのない行は無視する
                        td_elements = notice.find_elements(By.TAG_NAME,"td")
                        if not td_elements :
                            continue

                        target = notice.find_element(By.XPATH, ".//td[1]").text
                        print(f"対象: {target}")
                        update_time = notice.find_element(By.XPATH, ".//td[4]").text # 更新日時取得(文字列)

                        update_time_contain_year = f"{current_year}/{update_time}"
                        update_time_full = datetime.strptime(update_time_contain_year,"%Y/%m/%d %H:%M")
                        print(f"更新日時は{update_time_full}現在時刻は{now}")
                        
                        # 更新時刻が通知対象期間内であるか確認
                        if check_start_time <= update_time_full and update_time_full <= now:
                            print("更新時刻が通知対象期間内です")
                            #対象のキーワードが含まれていて、撮影済みでないか確認
                            if any (keyword in target for keyword in keywords) and i not in k :

                                    link = notice.find_element(By.XPATH, ".//td[5]/a")
                                    link.click()
                                    print("通知をクリックしました")
                                    
                                    #通知内容を取得
                                    element_text = driver.find_element(By.XPATH,"//table[contains(.,'おしらせ')]").text
                                    if ("2023/12/05 15:43:09" in element_text):
                                        if os.path.exists(healthcenter_text) :
                                            with open(healthcenter_text,mode="r",encoding="utf-8") as t :
                                                past_text = t.read()
                                            if (past_text == element_text):#もし変わってなければ通知しない。
                                                    driver.back()  # 前のページに戻る
                                                    print("ホームに戻りました")
                                                    # ページが読み込まれるのを待つ
                                                    WebDriverWait(driver, 10).until(
                                                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                                                    )
                                                    k.append(i)
                                                    break
                                        #変わってたら記録を更新する
                                        with open(healthcenter_text,mode="w",encoding="utf-8") as w : 
                                            w.write(element_text)
                                    # スクリーンショットを撮影
                                    screenshot_path = f"portal_screenshot_{i}.png"
                                    # ページが完全に読み込まれるのを待つ
                                    WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                                    )
                                    # ページの幅と高さを取得
                                    width = driver.execute_script("return document.body.scrollWidth")
                                    height = driver.execute_script("return document.body.scrollHeight")
                                    
                                    # ウィンドウサイズをページのサイズに設定
                                    driver.set_window_size(width, height)
                                    # スクリーンショットを撮影
                                    png = driver.find_element(By.XPATH,"//table[contains(.,'おしらせ')]").screenshot_as_png
                                    with open(screenshot_path,"wb") as f :
                                        f.write(png)

                                    print(f"スクリーンショットを撮影しました: {screenshot_path}")
                                    
                                    # LINE Notifyに画像を送信
                                    url = "https://notify-api.line.me/api/notify"
                                    if "個別" in target :
                                        LINE_Notify(LINE_NOTIFY_TOKEN_MINE,"ポータルが更新されました",screenshot_path)
                                        shot += 2
                                    else :
                                        LINE_Notify(LINE_NOTIFY_TOKEN,"ポータルが更新されました",screenshot_path)
                                        shot = 1
                                    print("LINEnotifyに画像を送信しました")
                                    
                                    # 撮影したスクリーンショットをデバイスから削除
                                    if os.path.exists(screenshot_path) :
                                        os.remove(screenshot_path)

                                    driver.back()  # 前のページに戻る
                                    print("ホームに戻りました")
                                    # ページが読み込まれるのを待つ
                                    WebDriverWait(driver, 10).until(
                                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                                    )

                                    k.append(i) # 撮影したもののインデックスを追加

                                    #time.sleep(1) # LINEnotifyの対応を待つ

                                    break

                            elif i in k :
                                print("撮影済みです")
                            
                            else :
                                print("対象外です")
                            
                        else :
                            print("更新時刻が通知対象期間外です")
                            j = 1 # 更新時刻が通知対象期間外のお知らせは確認しない
                            break
                        if i == len(notices):
                            j = 1
                            break # 最後まで確認したらbreak

                    # うまくいかなかったらその行をパスする
                    except (NoSuchElementException):
                        print("対象要素が見つかりません")
                        slack_notify("対象要素が見つからなかった…",1)
                        error = 1
                        break
                    except (StaleElementReferenceException):
                        print("対象要素が期限切れです")
                        slack_notify("対象要素がなくなっちゃった…",1)
                        error = 1
                        break
            j = 0 # 他のテーブル用にループを再開
            k.clear() # 撮影済みリストをクリア
            print(f"Table{table_name}を確認しました。")

    except Exception as e :
        slack_notify(f"{e.__class__.__name__}:{e}",1)
        print(f"{e.__class__.__name__}:{e}")
        error = 1

    finally:
        driver.quit()
        print("ブラウザを閉じました")
        
        if shot == 0 and error == 0 :
            slack_notify("なかったよ！")
            print ("ポータル通知がないことをSlackに送信しました。")
        
        if shot == 1 or shot == 3 :
            # 実行時刻を記録
            with open(time_record_file,mode="w",encoding="utf-8") as f : 
                f.write(donetime)
            slack_notify("あったから送っておいた！")
            print ("ポータルを確認したことをSlackに送信しました。")

        elif shot == 2 :
            # 自分にだけ送信したとき
            # 実行時刻を記録
            with open(time_record_file,mode="w",encoding="utf-8") as f : 
                f.write(donetime)
            slack_notify("と・く・べ・つ♡",1)
            print ("自分の分だけ通知があったことをSlackに送信しました。")
try:
    if len(os.listdir(remote_switch_path)) :
        print("remote_switchにファイルが存在するため、実行します。")
        # メインの関数を実行
        take_screenshot_and_send()
    else :
        slack_notify("わたしに仕事をさせてくれないのね…")
        print("remote_switchにファイルが存在しないため、実行しませんでした。")
except Exception as e :
    if str(e) == r"[WinError 3] 指定されたパスが見つかりません。: '\Portaのスイッチ'":
        slack_notify("仕事していいのかわからないよ…",1)
    else:
        slack_notify(f"{e.__class__.__name__}:{e}",1)
        print(f"{e.__class__.__name__}:{e}")