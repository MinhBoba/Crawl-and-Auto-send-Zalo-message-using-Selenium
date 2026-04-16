import pandas as pd
import random
import time
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ================= CONFIGURATION =================
EXCEL_FILE = 'CongTy_HCM.xlsx'
ZALO_COLUMN_NAME = 'Zalo'

MESSAGES = [
    "Em chào anh/chị ạ, bên mình có nhận gửi hàng 16kg từ HCM đi Quảng Châu giao tận nhà k ạ? Cho e xin giá với ạ.",
    "Cho em hỏi bên mình có hỗ trợ ship hàng 16kg từ HCM đi Quảng Châu giao tận nhà k ạ? Cho e xin giá với ạ.",
    "Em cần gửi hàng từ TP.HCM đi Quảng Châu khoảng 16kg giao tận nhà, anh/chị báo giá giúp em với ạ.",
    "Bên mình có dịch vụ vận chuyển từ HCM đi Quảng Châu giao tận nhà k ạ? Em cần gửi khoảng 16kg. Cho e xin giá với ạ.",
]
# =================================================


def human_typing(element, text):
    """Simulate human typing"""
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.03, 0.1))


def extract_phone_number(zalo_link):
    """Extract phone number from zalo.me link"""
    match = re.search(r'zalo\.me/([0-9a-zA-Z\+]+)', zalo_link)
    return match.group(1) if match else None


def main():
    # 1. Load Excel file
    try:
        df = pd.read_excel(EXCEL_FILE)
        print(f"Loaded {len(df)} contacts from Excel.")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # 2. Setup Chrome
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)

    # 3. Open Zalo Web
    driver.get("https://chat.zalo.me/")
    print("\n" + "=" * 50)
    print("Please scan QR code to log in to Zalo Web.")
    input("After login is complete, press ENTER to start...")
    print("=" * 50 + "\n")

    wait = WebDriverWait(driver, 10)

    # 4. Loop through contacts
    for index, row in df.iterrows():
        zalo_link = str(row.get(ZALO_COLUMN_NAME, ''))

        phone_number = extract_phone_number(zalo_link)
        if not phone_number:
            if zalo_link.strip() != 'nan':
                print(f"[{index+1}] Invalid link: {zalo_link}")
            continue

        chat_url = f"https://chat.zalo.me/?phone={phone_number}"
        print(f"[{index+1}/{len(df)}] Opening chat: {phone_number}")

        try:
            # Open chat in SAME TAB (important)
            driver.get(chat_url)

            # Wait before interaction
            time.sleep(random.uniform(3, 6))

            # Click "Message" button if needed
            try:
                message_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@data-translate-inner='STR_CHAT']"))
                )
                message_btn.click()
                print("   -> Clicked 'Message'")
                time.sleep(random.uniform(1, 2))
            except TimeoutException:
                pass  # Already inside chat

            # Find chat input
            chat_box = wait.until(
                EC.presence_of_element_located((By.ID, "richInput"))
            )

            chat_box.click()
            time.sleep(random.uniform(0.5, 1.2))

            # Pick random Vietnamese message
            message_text = random.choice(MESSAGES)

            # Type message (NOT sending)
            human_typing(chat_box, message_text)

            print("   => Message typed (not sent)")

            # Wait before next contact
            time.sleep(random.uniform(8, 15))

        except Exception as e:
            print(f"   => Error with {phone_number}: {e}")

        # Delay between contacts
        time.sleep(random.uniform(10, 25))

    print("\nFinished processing all contacts.")
    # driver.quit()


if __name__ == "__main__":
    main()