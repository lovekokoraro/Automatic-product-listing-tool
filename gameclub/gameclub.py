import random
import tkinter as tk
from tkinter import ttk, scrolledtext
from threading import Thread
from playwright.sync_api import sync_playwright
import time
import csv
import os
from datetime import datetime

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

class RMTScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("自動出品システム")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        self.is_scraping = False
        
        # Create UI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Title
        title_label = tk.Label(
            self.root, 
            text="gameclub自動出品システム", 
            font=("Noto Sans JP", 16, "bold"),
            pady=10
        )
        title_label.pack()
        
        # Status label
        self.status_label = tk.Label(
            self.root,
            text="状態: 準備完了",
            font=("Noto Sans JP", 10),
            fg="green"
        )
        self.status_label.pack(pady=5)
        
        # Login info frame
        login_frame = tk.Frame(self.root)
        login_frame.pack(pady=10)
        
        # Email input field
        email_label = tk.Label(login_frame, text="メール:", font=("Noto Sans JP", 10))
        email_label.pack(side=tk.LEFT, padx=5)
        self.email_entry = tk.Entry(login_frame, width=25, font=("Noto Sans JP", 10))
        self.email_entry.insert(0, "")
        self.email_entry.pack(side=tk.LEFT, padx=5)
        
        # Password input field
        password_label = tk.Label(login_frame, text="パスワード:", font=("Noto Sans JP", 10))
        password_label.pack(side=tk.LEFT, padx=5)
        self.password_entry = tk.Entry(login_frame, width=15, font=("Noto Sans JP", 10), show="*")
        self.password_entry.insert(0, "")
        self.password_entry.pack(side=tk.LEFT, padx=5)
        
        # Buttons and input fields frame
        buttons_frame = tk.Frame(self.root)
        buttons_frame.pack(pady=10)
        
        # Start button
        self.start_button = tk.Button(
            buttons_frame,
            text="出品開始",
            command=self.start_scraping,
            font=("Noto Sans JP", 12),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
            cursor="hand2"
        )
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        # Stop button (initially disabled)
        self.stop_button = tk.Button(
            buttons_frame,
            text="出品停止",
            command=self.stop_scraping,
            font=("Noto Sans JP", 12),
            bg="#FF6B35",
            fg="white",
            padx=20,
            pady=10,
            state="disabled",
            cursor="hand2"
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Input fields frame for st and ed (to the right of buttons)
        input_frame = tk.Frame(buttons_frame)
        input_frame.pack(side=tk.LEFT, padx=20)
        
        # st input field
        st_label = tk.Label(input_frame, text="下限:", font=("Noto Sans JP", 10))
        st_label.pack(side=tk.LEFT, padx=5)
        self.st_entry = tk.Entry(input_frame, width=10, font=("Noto Sans JP", 10))
        self.st_entry.insert(0, "300")
        self.st_entry.pack(side=tk.LEFT, padx=5)
        
        # ed input field
        ed_label = tk.Label(input_frame, text="上限:", font=("Noto Sans JP", 10))
        ed_label.pack(side=tk.LEFT, padx=5)
        self.ed_entry = tk.Entry(input_frame, width=10, font=("Noto Sans JP", 10))
        self.ed_entry.insert(0, "600")
        self.ed_entry.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            self.root,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=10)
        
        # Log text area
        log_label = tk.Label(
            self.root,
            text="ログ:",
            font=("Noto Sans JP", 10, "bold"),
            anchor="w"
        )
        log_label.pack(pady=(10, 5), padx=20, anchor="w")
        
        self.log_text = scrolledtext.ScrolledText(
            self.root,
            height=15,
            width=80,
            font=("Consolas", 9),
            wrap=tk.WORD
        )
        self.log_text.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
    def log(self, message):
        """Add message to log text area"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=f"状態: {message}", fg=color)
        
    def start_scraping(self):
        """Start scraping in a separate thread"""
        if not self.is_scraping:
            self.is_scraping = True
            self.start_button.config(state="disabled")
            self.stop_button.config(state="normal")
            self.progress.start()
            self.update_status("出品中...", "blue")
            self.log_text.delete(1.0, tk.END)
            self.log("出品処理を開始しています...")
            
            # Run scraping in a separate thread
            thread = Thread(target=self.scrape_gameclub, daemon=True)
            thread.start()
            
    def stop_scraping(self):
        """Stop scraping"""
        self.is_scraping = False
        self.log("出品を停止しています...")
        self.update_status("停止", "red")
        self.progress.stop()
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

        
    def scrape_gameclub(self):
        """Main scraping function"""
        # Get st and ed values from input fields
        try:
            st = int(self.st_entry.get())
            ed = int(self.ed_entry.get())
        except ValueError:
            self.log("無効な下限または上限の値です。デフォルト値（300、600）を使用します")
            st = 300
            ed = 600
        
        def get_image_files(id, file_type="image"):
            folder = f"./images/{id}"
            if not os.path.exists(folder):
                return []
            try:
                if file_type == "image":
                    return [
                        os.path.abspath(os.path.join(folder, f))
                        for f in os.listdir(folder)
                        if f.lower().endswith((".jpg", ".jpeg", ".png", ".gif"))
                    ]
                elif file_type == "video":
                    return [
                        os.path.abspath(os.path.join(folder, f))
                        for f in os.listdir(folder)
                        if f.lower().endswith((".mp4", ".mov", ".avi", ".mkv", ".mpg", ".mpeg", ".webm"))
                    ]
            except Exception:
                return []
        def waiting (st,ed):
            sleep_time = random.randint(st, ed)
            elapsed = 0
            while elapsed < sleep_time and self.is_scraping:
                self.log(f"waiting {elapsed} seconds")
                time.sleep(1)
                elapsed += 1
                if not self.is_scraping:
                    return
        try:
            # Open and read rmt.csv
            with sync_playwright() as p:
                self.log("ブラウザを起動しています...")
                try:
                    browser = p.chromium.launch(headless=False)
                except Exception as browser_error:
                    if "Executable doesn't exist" in str(browser_error) or "browser" in str(browser_error).lower():
                        self.log("エラー: Playwrightブラウザがインストールされていません！")
                        self.log("次のコマンドを実行してブラウザをインストールしてください：")
                        self.log(" playwright install chromium")
                        self.update_status("エラー: ブラウザがインストールされていません", "red")
                        return
                    raise
                
                if not self.is_scraping:
                    browser.close()
                    return
                
                self.log("新しいページを作成しています...")
                page = browser.new_page()
                
                self.log("https://gameclub.jp/signin に移動しています...")
                page.goto("https://gameclub.jp/signin")
                self.log("ページの読み込みを待っています...")
                page.wait_for_load_state("load")
                self.log("ページの読み込みが完了しました！")
                # Get login credentials from input fields
                email = self.email_entry.get()
                password = self.password_entry.get()
                page.fill("input[name='email']", email)
                page.fill("input[name='password']", password)
                
                # Wait for the button with class "btn_type1 fade" to be available
                self.log("ログインを待っています...")
                # page.wait_for_selector(".btn_type1.fade", state="detached", timeout=1000000)
                while True:
                    time.sleep(1)
                    try:
                        exist = page.locator("input[name='password']").count()
                        if not self.is_scraping:
                            browser.close()
                            return
                        if exist == 0:
                            break
                    except Exception as e:
                        continue
                
                
                self.log("products.csvを開いています...")
                # Read all rows from CSV
                all_rows = []
                with open("products.csv", "r", newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    all_rows = list(reader)
                
                is_first = True
                for id, row in enumerate(all_rows):
                    if id == 0:
                        continue
                    if not self.is_scraping:
                        break
                    if row[12] == "ok":
                        self.log(f"商品{id}はすでに出品されています")
                        continue
                    
                    product_formtitle = row[0]
                    product_category = row[1]
                    product_title = row[2]
                    product_description = row[3]
                    product_billing = row[4]
                    product_comment = row[5]
                    product_notification = row[6]
                    product_method = row[7]
                    product_price = row[8]
                    product_discussion = row[9]
                    product_password = row[10]
                    product_uniqueness = row[11]

                    is_duplicate = False
                    for i in range(id - 1):
                        sum = 0
                        for j in range(12):
                            if all_rows[i][j] == row[j]:
                                sum += 1
                        if sum == 12:
                            is_duplicate = True
                            self.log(f"商品{id}はすでに出品されています")
                            continue
                    if is_duplicate:
                        self.log(f"商品{id}はすでに出品されています")
                        continue
                    
                    try:
                        page.goto("https://gameclub.jp/mypage/products/add")
                        page.wait_for_load_state("load")
                        if is_first:
                            is_first = False
                        else:
                            waiting(st, ed)
                        
                        if not self.is_scraping:
                            break
                        
                        
                        
                        #image upload
                        image_files = get_image_files(id,"image")
                        if not image_files:
                            self.log(f"商品{id}の画像フォルダが見つかりませんが、画像なしで出品を続行します...")
                        else:
                            page.query_selector('#image-upload-label').set_input_files(image_files)
                            waiting(3, 5)
                        
                        #upload video
                        video_files = get_image_files(id, "video")
                        if not video_files:
                            self.log(f"商品{id}の動画フォルダが見つかりませんが、動画なしで出品を続行します...")
                        else:
                            # Only use the first video file since the input doesn't support multiple files
                            page.query_selector('#movie-upload-label').set_input_files(video_files[0])
                            waiting(3, 5)
                        
                        if not self.is_scraping:
                            break
                        
                        #ゲームタイトル
                        self.log(f"product_formtitle: {product_formtitle}")
                        page.locator("input.btn-search-title").click()
                        page.wait_for_load_state("load")
                        waiting(3,3)
                        if not self.is_scraping:
                            break
                        page.locator('span.title-name', has_text=product_formtitle).click()
                        waiting(3,3)
                        
                        #カテゴリ
                        self.log(f"product_category: {product_category}")
                        page.locator('label', has_text=product_category)\
                            .locator('input')\
                            .click()

                        #出品タイトル
                        self.log(f"product_title: {product_title}")
                        page.locator('#name').fill(product_title)
                        
                        #商品説明
                        self.log(f"product_description: {product_description}")
                        page.locator('#input-body-text').fill(product_description)
                        
                        #課金総額
                        self.log(f"product_billing: {product_billing}")
                        page.locator('input[name="subcategory_unique_property_1_value"]').fill(product_billing)
                        
                        #買い手へ初回自動表示するメッセージ
                        self.log(f"product_comment: {product_comment}")
                        page.locator('#firstchat').fill(product_comment)
                        
                        #出品を通知
                        self.log(f"product_notification: {product_notification}")
                        page.locator('input[name="notify_user_id"]').fill(product_notification)
                        
                        #出品方法
                        try:
                            self.log(f"product_method: {product_method}")
                            page.locator(f"label:has-text('{product_method}')").click()
                        except Exception as e:
                            self.log(f"product_method: {product_method} が見つかりませんでした")
                        
                        #商品価格
                        try:
                            self.log(f"product_price: {product_price}")
                            page.locator('input[name="price"]').fill(product_price)
                        except Exception as e:
                            self.log(f"product_price: {product_price} が見つかりませんでした")
                        
                        #金額の相談
                        try:
                            self.log(f"product_discussion: {product_discussion}")
                            if product_discussion == "1":
                                page.locator(f"label:has-text('金額の相談')").click()
                        except Exception as e:
                            self.log(f"product_discussion: {product_discussion} が見つかりませんでした")
                        
                        #暗証番号を設定する
                        try:
                            self.log(f"product_password: {product_password}")
                            if not product_password == "":
                                page.locator("label:has-text('暗証番号を設定する')").click()
                                page.locator('#pin').fill(product_password)
                        except Exception as e:
                            self.log(f"product_password: {product_password} が見つかりませんでした")
                            
                        #この商品は「他サイトでは出品していない、ゲームクラブだけで出品している」商品ですか？
                        try:
                            if product_uniqueness == "はい":
                                page.locator('input[name="is_gc_only"][value="1"]').check()
                            elif product_uniqueness == "いいえ":
                                page.locator('input[name="is_gc_only"][value="0"]').check()                                
                        except Exception as e:
                            self.log(f"product_uniqueness: {product_uniqueness} が見つかりませんでした")
                        
                        waiting(3,3)
                        if not self.is_scraping:
                            break
                        
                        page.locator("#btn-confirm").click()
                        page.wait_for_load_state("networkidle")
                        self.log("btn-confirmをクリックしました")
                        waiting(3,3)
                        if not self.is_scraping:
                            break
                        
                        page.locator("#btn-add").click()
                        page.wait_for_load_state("networkidle")
                        page.wait_for_load_state("load")
                        self.log("btn-addをクリックしました")
                        
                        waiting(3,3)
                        current_url = page.url
                        self.log(f"current_url: {current_url}")
                        # Product successfully exhibited - update CSV
                        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        row[12] = "ok"
                        row[13] = current_time
                        all_rows[id] = row
                        
                        # Write updated CSV immediately
                        with open("products.csv", "w", newline="", encoding="utf-8") as f:
                            writer = csv.writer(f)
                            writer.writerows(all_rows)
                        
                        self.log(f"商品{id}の出品が成功しました！（CSVに'ok'とマークされました）")
                        time.sleep(10)
                        
                    except Exception as e:
                        self.log(f"商品{id}の処理中にエラーが発生しました：{str(e)}")
                        # Don't mark as "ok" if there was an error
                        continue
                
                if self.is_scraping:
                    self.log("すべての商品の処理が完了しました！")
                    self.update_status("完了", "green")
                else:
                    self.log("ユーザーによって停止されました。")
                
        except Exception as e:
            self.log(f"エラーが発生しました：{str(e)}")
            self.update_status("エラー", "red")
        finally:
            self.is_scraping = False
            self.progress.stop()
            self.start_button.config(state="normal")
            self.stop_button.config(state="disabled")

def main():
    root = tk.Tk()
    app = RMTScraperGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

