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
        self.root.title("RMT自動出品システム")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        self.is_scraping = False
        
        # Create UI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Title
        title_label = tk.Label(
            self.root, 
            text="RMT自動出品システム", 
            font=("Arial", 16, "bold"),
            pady=10
        )
        title_label.pack()
        
        # Status label
        self.status_label = tk.Label(
            self.root,
            text="状態: 準備完了",
            font=("Arial", 10),
            fg="green"
        )
        self.status_label.pack(pady=5)
        
        # Login info frame
        login_frame = tk.Frame(self.root)
        login_frame.pack(pady=10)
        
        # Email input field
        email_label = tk.Label(login_frame, text="メール:", font=("Arial", 10))
        email_label.pack(side=tk.LEFT, padx=5)
        self.email_entry = tk.Entry(login_frame, width=25, font=("Arial", 10))
        self.email_entry.insert(0, "")
        self.email_entry.pack(side=tk.LEFT, padx=5)
        
        # Password input field
        password_label = tk.Label(login_frame, text="パスワード:", font=("Arial", 10))
        password_label.pack(side=tk.LEFT, padx=5)
        self.password_entry = tk.Entry(login_frame, width=15, font=("Arial", 10), show="*")
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
            font=("Arial", 12),
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
            font=("Arial", 12),
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
        st_label = tk.Label(input_frame, text="下限:", font=("Arial", 10))
        st_label.pack(side=tk.LEFT, padx=5)
        self.st_entry = tk.Entry(input_frame, width=10, font=("Arial", 10))
        self.st_entry.insert(0, "300")
        self.st_entry.pack(side=tk.LEFT, padx=5)
        
        # ed input field
        ed_label = tk.Label(input_frame, text="上限:", font=("Arial", 10))
        ed_label.pack(side=tk.LEFT, padx=5)
        self.ed_entry = tk.Entry(input_frame, width=10, font=("Arial", 10))
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
            font=("Arial", 10, "bold"),
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
            thread = Thread(target=self.scrape_rmt_club, daemon=True)
            thread.start()
            
    def stop_scraping(self):
        """Stop scraping"""
        self.is_scraping = False
        self.log("出品を停止しています...")
        self.update_status("停止", "red")
        self.progress.stop()
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")

        
    def scrape_rmt_club(self):
        """Main scraping function"""
        # Get st and ed values from input fields
        try:
            st = int(self.st_entry.get())
            ed = int(self.ed_entry.get())
        except ValueError:
            self.log("無効な下限または上限の値です。デフォルト値（300、600）を使用します")
            st = 300
            ed = 600
        
        def get_image_files(id):
            folder = f"./images/{id}"
            if not os.path.exists(folder):
                return []
            try:
                return [
                    os.path.abspath(os.path.join(folder, f))
                    for f in os.listdir(folder)
                    if f.lower().endswith((".jpg", ".jpeg", ".png", ".gif"))
                ]
            except Exception:
                return []
        def waiting (st,ed):
            sleep_time = random.randint(st, ed)
            elapsed = 0
            while elapsed < sleep_time and self.is_scraping:
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
                        self.log("  playwright install chromium")
                        self.update_status("エラー: ブラウザがインストールされていません", "red")
                        return
                    raise
                
                if not self.is_scraping:
                    browser.close()
                    return
                
                self.log("新しいページを作成しています...")
                page = browser.new_page()
                
                self.log("https://rmt.club/user-login に移動しています...")
                page.goto("https://rmt.club/user-login")
                self.log("ページの読み込みを待っています...")
                page.wait_for_load_state("load")
                self.log("ページの読み込みが完了しました！")
                # Get login credentials from input fields
                email = self.email_entry.get()
                password = self.password_entry.get()
                page.fill("input[name='data[User][mail]']", email)
                page.fill("input[name='data[User][password]']", password)
                
                # Wait for the button with class "btn_type1 fade" to be available
                self.log("ログインを待っています...")
                # page.wait_for_selector(".btn_type1.fade", state="detached", timeout=1000000)
                while True:
                    try:
                        exist = page.query_selector(".btn_type1.fade")
                        if not self.is_scraping:
                            browser.close()
                            return
                        if not exist:
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
                    
                    if row[6] == "ok":
                        self.log(f"商品{id}の出品が成功しました！（CSVに'ok'とマークされました）")
                        continue
                    
                    is_duplicate = False
                    for i in range(id - 1):
                        sum = 0
                        for j in range(6):
                            if all_rows[i][j] == row[j]:
                                sum += 1
                        if sum == 6:
                            is_duplicate = True
                            self.log(f"商品{id}はすでに出品されています")
                            continue
                    if is_duplicate:
                        self.log(f"商品{id}はすでに出品されています")
                        continue
                    
                    
                    product_name = row[0]
                    product_title = row[1]
                    product_tag = row[2]
                    product_detail = row[3]
                    product_notification = row[4]
                    product_price = row[5]
                    self.log(f"product_name: {product_name}")
                    self.log(f"product_title: {product_title}")
                    self.log(f"product_tag: {product_tag}")
                    self.log(f"product_detail: {product_detail}")
                    self.log(f"product_notification: {product_notification}")
                    self.log(f"product_price: {product_price}")
                
                    try:
                        page.goto("https://rmt.club/deals/add")
                        page.wait_for_load_state("load")
                        page.click("span.deal_type")
                        if is_first:
                            is_first = False
                            waiting(10, 15)
                        else:
                            waiting(st, ed)
                        
                        if not self.is_scraping:
                            break
                        
                        page.fill("input[name='data[Deal][game_title]']", product_name)
                        page.fill("input[name='data[Deal][deal_title]']", product_title)
                        page.fill("input[name='data[Deal][tag]']", product_tag)
                        page.fill("textarea[name='data[Deal][info]']", product_detail)
                        page.fill("input[name='data[Deal][user_name]']", product_notification)
                        page.fill("input[name='data[Deal][deal_price]']", product_price)
                        
                        image_files = get_image_files(id)
                        if not image_files:
                            self.log(f"商品{id+1}の画像フォルダが見つかりませんが、画像なしで出品を続行します...")
                        else:
                            page.query_selector('input[type="file"][name="data[Deal][upload_img][]"]').set_input_files(image_files)
                            waiting(8, 10)
                        
                        if not self.is_scraping:
                            break
                        
                        page.query_selector("input[name='smt_confirm']").click()
                        page.wait_for_load_state("load")
                        waiting(5, 10)
                        
                        if not self.is_scraping:
                            break
                        
                        page.query_selector('input[name="data[Deal][agreement]"]').check();
                        page.click(".btn_type1.fade.btn_next")
                        waiting(5, 10)
                        
                        # Product successfully exhibited - update CSV
                        row[6] = "ok"
                        row[7] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        all_rows[id] = row
                        
                        # Write updated CSV immediately
                        with open("products.csv", "w", newline="", encoding="utf-8") as f:
                            writer = csv.writer(f)
                            writer.writerows(all_rows)
                        
                        self.log(f"商品{id+1}の出品が成功しました！（CSVに'ok'とマークされました）")
                        
                    except Exception as e:
                        self.log(f"商品{id+1}の処理中にエラーが発生しました：{str(e)}")
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

