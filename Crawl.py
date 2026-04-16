from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time


def clean_phone(phone):
    """Clean and format the phone number"""
    phone = phone.replace('.', '').replace(' ', '').replace('-', '').strip()

    if phone.startswith('+84'):
        phone = '0' + phone[3:]

    if phone.isdigit() and phone.startswith('0') and len(phone) == 10:
        return phone

    return None


def extract_phones_and_address(block, next_block):
    phones =[]
    address = ""

    curr = block.next_element

    while curr and curr != next_block:
        if hasattr(curr, 'name') and curr.name:

            # Extract phone numbers
            if curr.name == 'a':
                href = curr.get('href', '')
                if href.startswith('tel:'):
                    phone = curr.text.strip()
                    if phone and phone not in phones:
                        phones.append(phone)

            # Extract the address
            if curr.name == 'i':
                classes = curr.get('class',[])
                if 'fa-location-dot' in classes:
                    if curr.parent:
                        address = curr.parent.text.strip()

        curr = curr.next_element

    return phones, address


def setup_driver():
    options = Options()
    # options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    )

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def crawl():
    driver = setup_driver()

    categories =[
        "https://trangvangvietnam.com/categories/419935/chuyen-phat-nhanh-cong-ty-chuyen-phat-nhanh.html",
        "https://trangvangvietnam.com/categories/484645/logistics-dich-vu-logistics.html"
    ]

    results =[]

    try:
        for base_url in categories:
            print(f"\n=== Crawling: {base_url} ===")

            page = 1
            prev_names =[]

            while True:
                url = f"{base_url}?page={page}"
                print(f"Page {page}")

                driver.get(url)
                time.sleep(2)

                soup = BeautifulSoup(driver.page_source, 'html.parser')
                blocks = soup.find_all('div', class_='listings_center')

                if not blocks:
                    print("No more pages found. Moving to the next category.")
                    break

                current_names =[]

                for i, block in enumerate(blocks):
                    name_tag = block.find('h2')
                    if not name_tag:
                        continue

                    name = name_tag.text.strip()
                    current_names.append(name)

                    next_block = blocks[i + 1] if i + 1 < len(blocks) else None

                    raw_phones, address = extract_phones_and_address(block, next_block)

                    # Clean extracted phone numbers
                    clean_phones =[]
                    for p in raw_phones:
                        if '(' in p or ')' in p:
                            continue

                        cp = clean_phone(p)
                        if cp and cp not in clean_phones:
                            clean_phones.append(cp)

                    # Filter for Ho Chi Minh City addresses only
                    address_lower = address.lower()
                    if "hồ chí minh" not in address_lower and "hcm" not in address_lower:
                        continue

                    phone_str = " - ".join(clean_phones) if clean_phones else "N/A"
                    zalo = f"https://zalo.me/{clean_phones[0]}" if clean_phones else ""

                    results.append({
                        "Tên Công Ty": name,
                        "SĐT": phone_str,
                        "Zalo": zalo,
                        "Địa Chỉ": address
                    })

                # Check for duplicate pages (to stop infinite loops if the site repeats the last page)
                if current_names == prev_names:
                    print("Reached the last page (duplicate content detected).")
                    break

                prev_names = current_names
                page += 1

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        driver.quit()

    # Export results to an Excel file
    if results:
        df = pd.DataFrame(results)
        df.drop_duplicates(subset=['Tên Công Ty'], inplace=True)

        df.to_excel("CongTy_HCM.xlsx", index=False)

        print(f"\nDONE: Successfully scraped {len(df)} companies.")
    else:
        print("No data collected.")


if __name__ == "__main__":
    crawl()