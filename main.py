import os
import argparse
import re
import json
from datetime import datetime
from pypdf import PdfReader  # ✨ 記得先 pip install pypdf
from playwright.sync_api import sync_playwright
from append_to_excel import create_new_sheet_and_fill

def parse_uber_date(date_str):
    current_year = datetime.now().year
    cleaned_str = date_str.strip()
    months_map = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
    }
    year_match = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", cleaned_str)
    if year_match:
        return datetime(int(year_match.group(1)), int(year_match.group(2)), int(year_match.group(3)))
    month_day_match = re.search(r"(\d{1,2})月\s*(\d{1,2})日", cleaned_str)
    if month_day_match:
        return datetime(current_year, int(month_day_match.group(1)), int(month_day_match.group(2)))
    eng_match = re.search(r"([A-Za-z]{3,})\s*(\d{1,2}),\s*(\d{4})", cleaned_str)
    if eng_match:
        return datetime(int(eng_match.group(3)), months_map.get(eng_match.group(1).lower()[:3], 1), int(eng_match.group(2)))
    return datetime.now()

# ========================================================
# ⚡ 【離線解析神醫】PDF 儲存後，立刻開箱讀取文字明細
# ========================================================
def extract_data_from_pdf(pdf_path, backup_timestamp):
    try:
        reader = PdfReader(pdf_path)
        pdf_text = ""
        for page in reader.pages:
            pdf_text += page.extract_text() + "\n"
        
        address = "未知地址"
        total_amount = "未知金額"
        items_list = []
        
        # 1. 提取總金額 (尋找 Total / 總計 欄位)
        fare_match = re.search(r"(?:Total|總計|Total fare|總金額)[:\s]*(\$?[\d.,]+)", pdf_text, re.IGNORECASE)
        if fare_match:
            total_amount = fare_match.group(1)
        else:
            all_prices = re.findall(r"\$[\d.,]+", pdf_text)
            if all_prices:
                total_amount = all_prices[-1]
        
        # 2. 提取外送地址
        # 2. 離線提取外送地址 (修正版：精準鎖定 Delivery 後方的地址)
        # 邏輯：尋找 Delivery / 外送至 之後，緊接著的地址特徵
        delivery_address_match = re.search(
            r"(?:Delivery|外送至|外送|Delivered)[^\n]*\n\s*([\d\s\w,.]+,\s*[A-Z]{2}\s*\d{5}[-\d]*)", 
            pdf_text, 
            re.IGNORECASE
        )
        
        if delivery_address_match:
            address = delivery_address_match.group(1).strip()
        else:
            # 保底機制：如果找不到帶有 Delivery 的地址，才退回抓取第一個標準地址
            address_match = re.search(r"([\d\s\w,.]+,\s*[A-Z]{2}\s*\d{5}[-\d]*)", pdf_text)
            if address_match:
                address = address_match.group(1).strip()
            else:
                # 台灣地址保底
                tw_match = re.search(r"([^ \n]*(?:市|縣)[^ \n]*(?:區|市|鎮|鄉)[^ \n]*路[^ \n]*)", pdf_text)
                if tw_match:
                    address = tw_match.group(1).strip()
                
        # 3. 提取餐點明細
        lines = pdf_text.split("\n")
        start_collecting = False
        for line in lines:
            cleaned_line = line.strip()
            if not cleaned_line:
                continue
            if any(k in cleaned_line for k in ["Items", "品項", "Order Detail", "訂單明細"]):
                start_collecting = True
                continue
            if any(k in cleaned_line for k in ["Subtotal", "小計", "Total", "總計", "Delivered to", "外送至"]):
                start_collecting = False
                
            if start_collecting:
                items_list.append({"item_raw_text": cleaned_line})
                
        return {
            "order_timestamp": backup_timestamp,
            "delivery_address": address,
            "total_amount": total_amount,
            "items_detail": items_list
        }
    except Exception as e:
        print(f"⚠️ 離線解析 PDF 失敗: {e}")
        return {"order_timestamp": backup_timestamp, "delivery_address": "解析失敗", "total_amount": "解析失敗", "items_detail": []}

def main(start_date_str, end_date_str, excel_file_name):
    print(f"📅 目標日期區間: {start_date_str} 至 {end_date_str}")
    print(f"📂 目標 Excel 檔案: {excel_file_name}")
    start_limit = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_limit = datetime.strptime(end_date_str, "%Y-%m-%d")
    
    output_dir = f"./uber_receipts_{start_date_str}_to_{end_date_str}"
    os.makedirs(output_dir, exist_ok=True)

    with sync_playwright() as p:
        user_data_dir = "./uber_chrome_profile"
        
        print("正在啟動瀏覽器...")
        context = p.chromium.launch_persistent_context(user_data_dir, headless=False, args=["--start-maximized"])
        page = context.new_page()
        
        print("前往 Uber Eats 訂單頁面...")
        page.goto("https://www.ubereats.com/tw/orders")
        
        if "login" in page.url or page.locator("text=登入").is_visible():
            print("⚠️ 請手動完成登入...")
            input()
            
        page.wait_for_selector("text=/過去的訂單|Past Orders/", timeout=60000)
        print(f"開始篩選區間 {start_date_str} 至 {end_date_str}...")
        
        order_selector = "div._al._am._bh"
        downloaded_orders = set()

        while True:
            orders = page.locator(order_selector).all()
            if not orders:
                break
                
            last_order_in_page_out_of_range = False

            for index, order in enumerate(orders):
                card_text = order.inner_text()
                order_fingerprint = card_text[:100].replace("\n", "")
                if order_fingerprint in downloaded_orders:
                    continue
                
                date_text = ""
                for line in card_text.split("\n"):
                    if "月" in line and "日" in line:
                        date_text = line.strip()
                        break
                if not date_text:
                    continue
                
                order_date = parse_uber_date(date_text)
                
                if order_date > end_limit:
                    downloaded_orders.add(order_fingerprint)
                    continue
                elif order_date < start_limit:
                    print(f"🛑 訂單日期已超出區間，停止搜尋。")
                    last_order_in_page_out_of_range = True
                    break
                
                # --------------------------------------------------------
                # 🛠️ 網頁端只做最穩定的點擊下載，其餘完全交給 Python 離線處理
                # --------------------------------------------------------
                receipt_btn = order.locator("a:has-text('View Receipt'), a:has-text('檢視電子明細'), a:has-text('查看收據')")
                
                if receipt_btn.count() > 0:
                    try:
                        print(f"🎯 符合區間！正在打開第 {index+1} 筆訂單的收據彈窗...")
                        receipt_btn.click()
                        
                        # 1. 確保下載 PDF 按鈕已載入（不開新分頁、不找 iframe，直接在 page 盲等文字按鈕）
                        download_btn = page.locator("a:has-text('Download PDF'), a:has-text('下載 PDF')")
                        download_btn.wait_for(state="visible", timeout=6000)
                        
                        # 2. 攔截並執行 PDF 下載
                        with page.expect_download() as download_info:
                            download_btn.click()
                            
                        download = download_info.value
                        formatted_order_date = order_date.strftime("%Y%m%d")
                        unique_id = datetime.now().strftime("%H%M%S")
                        
                        # 正式保存 PDF 檔案
                        pdf_filename = f"UberEats_{formatted_order_date}_{unique_id}_{index}.pdf"
                        final_pdf_path = os.path.join(output_dir, pdf_filename)
                        download.save_as(final_pdf_path)
                        print(f"💾 【1/2 網頁端成功】PDF 已順利儲存: {pdf_filename}")
                        
                        # 3. ✨ 【亮點新核心】檔案存好後，立刻離線讀取 PDF，榨出你要的資料
                        print("⚡ 正在發動離線 PDF 明細解析...")
                        order_data = extract_data_from_pdf(final_pdf_path, date_text.strip())
                        
                        # 4. 同步儲存成同名 JSON 檔案
                        json_filename = f"UberEats_{formatted_order_date}_{unique_id}_{index}.json"
                        with open(os.path.join(output_dir, json_filename), "w", encoding="utf-8") as f:
                            json.dump(order_data, f, ensure_ascii=False, indent=4)
                        print(f"📝 【2/2 解析成功】JSON 結構化明細已同步生成！")
                        print(f"   📍 地址: {order_data['delivery_address']} | 💰 總額: {order_data['total_amount']}")
                        
                        downloaded_orders.add(order_fingerprint)
                        
                    except Exception as e:
                        print(f"❌ 處理第 {index+1} 筆收據流程失敗: {e}")
                        
                    finally:
                        # 5. 🔥 【強力清空舞台機制】雙重保險關閉彈窗，防止擋住下一筆
                        try:
                            # 保險 A：按 Escape
                            page.keyboard.press("Escape")
                            # 保險 B：精準尋找網頁上的任何關閉按鈕，看見就點掉
                            close_btn = page.locator("button[aria-label='Close'], button[aria-label='關閉'], button:has-text('Close')")
                            if close_btn.count() > 0 and close_btn.first.is_visible():
                                close_btn.first.click()
                        except:
                            pass
                        # 留給網頁 1 秒鐘的關閉動畫時間
                        page.wait_for_timeout(1000)
                else:
                    downloaded_orders.add(order_fingerprint)

            if last_order_in_page_out_of_range:
                break
                
            # 清單加載更多
            show_more_btn = page.locator("button:has-text('Show more'), button:has-text('顯示更多')")
            if show_more_btn.count() > 0 and show_more_btn.is_visible():
                show_more_btn.scroll_into_view_if_needed()
                show_more_btn.click()
                page.wait_for_timeout(3000)
            else:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                page.wait_for_timeout(2000)
                if len(page.locator(order_selector).all()) == len(orders):
                    break
        
        context.close()
        print("🎉 網頁端 PDF 下載與離線 JSON 明細擷取已全部完成！")
        
        # ========================================================
        # 🎯 【最佳化路徑修復】自動將檔名補全為「絕對路徑」
        # ========================================================
        print("\n📊 偵測到網頁端已收工，現在啟動 Excel 智慧分頁串接引擎...")
        
        # 1. 自動補全 .xlsx 延伸檔名
        if not excel_file_name.lower().endswith(".xlsx"):
            excel_file_name += ".xlsx"
            
        # 2. 🎯 關鍵修復：取得 main.py 目前所在的「資料夾絕對路徑」
        # 這樣可以確保不論從哪裡啟動，路徑起點都是專案資料夾內部
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 3. 把「資料夾路徑」與「Excel 檔名」完美黏合，變成完整路徑！
        # 例如會變成：C:\Users\penny\ForFun\UberEatsreceipt\2026 Meal Expense.xlsx
        full_excel_path = os.path.join(current_dir, excel_file_name)
        
        # 4. 同理，資料夾路徑也可以用絕對路徑組裝，更安全
        dynamic_json_dir = os.path.join(current_dir, f"uber_receipts_{start_date_str}_to_{end_date_str}")
        
        print(f"📂 正在處理資料夾: {dynamic_json_dir}")
        print(f"📄 完整 Excel 路徑: {full_excel_path}")
        
        # 5. 🎯 正式傳入！此時傳過去的 full_excel_path 就是完美的絕對路徑了
        create_new_sheet_and_fill(dynamic_json_dir, full_excel_path)
        
        print("🎉 所有流程完全結束，您可以放心關閉此視窗或查看 Excel 成果囉！")



# ========================================================
# 🎯 修改：讓 Python 能接收 --excel 參數
# ========================================================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Uber Eats 收據自動下載工具")
    parser.add_argument("--start", required=True, help="開始日期 (YYYY-MM-DD)")
    parser.add_argument("--end", required=True, help="結束日期 (YYYY-MM-DD)")
    
    # 🎯 新增：接收 Excel 檔名的參數設定
    parser.add_argument("--excel", required=True, help="既有的 Excel 檔案名稱 (.xlsx)")
    
    args = parser.parse_args()
    
    main(args.start, args.end, args.excel)


