import requests
from pathlib import Path
import gzip
import pandas as pd
import jdatetime
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# تنظیمات اصلی
# ============================================================

URL = "https://old.tsetmc.com/tsev2/excel/MarketWatchPlus.aspx?d=0"

SAVE_FOLDER = Path(r"C:\MIC")

TARGET_SYMBOLS = [
    "فارماکیان",
    "دارونو",
    "فارمانی"
]


# ============================================================
# استانداردسازی متن
# ============================================================

def normalize_text(text):

    text = str(text)

    # یکسان سازی حروف عربی و فارسی
    text = text.replace("ك", "ک")
    text = text.replace("ي", "ی")
    text = text.replace("ى", "ی")

    # حذف فاصله ها
    text = text.replace(" ", "")
    text = text.replace("\u200c", "")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")
    text = text.replace("\u00a0", "")

    # حذف کاراکترهای اضافی
    text = text.replace("\t", "")
    text = text.replace("\r", "")
    text = text.replace("\n", "")

    return text.strip()


# ============================================================
# مرحله 1
# دانلود اطلاعات TSETMC
# ============================================================

def download_market_watch():

    print("\n")
    print("#" * 60)
    print("#        سیستم دریافت اطلاعات صندوق های دارویی")
    print("#" * 60)

    print("\n")
    print("=" * 60)
    print("مرحله 1 - دریافت اطلاعات از TSETMC")
    print("=" * 60)

    try:

        SAVE_FOLDER.mkdir(
            parents=True,
            exist_ok=True
        )

        print("\nدر حال اتصال به TSETMC...")
        print(URL)

        headers = {
            "User-Agent": (
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/153.0 Safari/537.36"
            ),
            "Accept": "*/*",
            "Referer": "https://www.tsetmc.com/"
        }

        response = requests.get(
            URL,
            headers=headers,
            timeout=60
        )

        print("\nکد پاسخ سرور:", response.status_code)

        if response.status_code != 200:

            print("❌ دریافت فایل ناموفق بود.")

            return None

        content = response.content

        print(
            "حجم اطلاعات دریافت شده:",
            f"{len(content):,}",
            "Bytes"
        )

        # بررسی GZIP
        if content[:2] == b"\x1f\x8b":

            print("فایل به صورت GZIP دریافت شده است.")
            print("در حال خارج کردن فشرده سازی...")

            content = gzip.decompress(content)

            print("فشرده سازی با موفقیت انجام شد.")

        if not content or len(content) < 100:

            print("❌ اطلاعات دریافت شده معتبر نیست.")

            return None

        # گرفتن تاریخ و ساعت لحظه ذخیره
        now = jdatetime.datetime.now()

        jalali_date = now.strftime("%Y-%m-%d")
        current_time = now.strftime("%H-%M-%S")

        file_name = (
            f"MarketWatchPlus_"
            f"{jalali_date}_"
            f"{current_time}.xlsx"
        )

        file_path = SAVE_FOLDER / file_name

        print("\nدر حال ذخیره فایل...")

        print(
            "تاریخ شمسی:",
            jalali_date
        )

        print(
            "ساعت:",
            current_time.replace("-", ":")
        )

        with open(file_path, "wb") as file:

            file.write(content)

        if file_path.exists() and file_path.stat().st_size > 0:

            file_size = file_path.stat().st_size

            print("\n" + "=" * 60)
            print("✅ دانلود و ذخیره فایل با موفقیت انجام شد")
            print("=" * 60)

            print("\nنام فایل:")
            print(file_path.name)

            print("\nتاریخ ثبت:")
            print(jalali_date.replace("-", "/"))

            print("\nساعت ثبت:")
            print(current_time.replace("-", ":"))

            print("\nمسیر کامل فایل:")
            print(file_path)

            print("\nحجم فایل:")
            print(f"{file_size:,} Bytes")

            print("=" * 60)

            return file_path

        print("❌ فایل ذخیره نشد.")

        return None

    except Exception as e:

        print("\n")
        print("=" * 60)
        print("❌ خطا در مرحله دانلود")
        print("=" * 60)

        print("\nشرح خطا:")
        print(e)

        print("=" * 60)

        return None


# ============================================================
# مرحله 2
# استخراج اطلاعات هر سه صندوق از Excel
# ============================================================

def find_all_symbols(file_path):

    print("\n")
    print("=" * 60)
    print("مرحله 2 - جستجوی سه صندوق در TSETMC")
    print("=" * 60)

    try:

        print("\nدر حال باز کردن فایل Excel...")

        df = pd.read_excel(
            file_path,
            sheet_name="دیده بان بازار",
            header=None
        )

        print(
            "\nتعداد ردیف ها:",
            len(df)
        )

        print(
            "تعداد ستون ها:",
            len(df.columns)
        )

        print("\n✅ فقط ستون A بررسی خواهد شد.")

        results = {}

        # ----------------------------------------------------
        # بررسی مستقل هر سه صندوق
        # ----------------------------------------------------

        for target_symbol in TARGET_SYMBOLS:

            normalized_target = normalize_text(
                target_symbol
            )

            found = False

            for index in range(3, len(df)):

                original_value = df.iloc[index, 0]

                normalized_value = normalize_text(
                    original_value
                )

                if normalized_value == normalized_target:

                    results[target_symbol] = {

                        "symbol": df.iloc[index, 0],

                        "name": df.iloc[index, 1],

                        "transactions": df.iloc[index, 2],

                        "volume": df.iloc[index, 3],

                        "value": df.iloc[index, 4],

                        "yesterday": df.iloc[index, 5],

                        "first": df.iloc[index, 6],

                        "last": df.iloc[index, 7],

                        "last_change": df.iloc[index, 8],

                        "last_percent": df.iloc[index, 9],

                        "closing": df.iloc[index, 10],

                        "closing_change": df.iloc[index, 11],

                        "closing_percent": df.iloc[index, 12]
                    }

                    found = True

                    break

            if not found:

                results[target_symbol] = None

        return results

    except Exception as e:

        print("\n")
        print("=" * 60)
        print("❌ خطا در خواندن Excel")
        print("=" * 60)

        print("\nشرح خطا:")
        print(e)

        print("=" * 60)

        return None


# ============================================================
# مرحله 3
# دریافت حباب و NAV از TradersArena
# ============================================================

def get_tradersarena_data():

    print("\n")
    print("=" * 60)
    print("مرحله 3 - دریافت اطلاعات TradersArena")
    print("=" * 60)

    # --------------------------------------------------------
    # تنظیمات Chrome
    # --------------------------------------------------------

    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    # --------------------------------------------------------
    # مسیر ChromeDriver
    # --------------------------------------------------------

    service = Service(
        r"C:\chrome-test\chrome-win64\chromedriver.exe"
    )

    driver = webdriver.Chrome(
        service=service,
        options=options
    )

    tradersarena_results = {}

    try:

        print("\nدر حال ورود به TradersArena...")

        driver.get(
            "https://tradersarena.ir/industries/sector-funds"
        )

        wait = WebDriverWait(
            driver,
            30
        )

        # ----------------------------------------------------
        # منتظر ماندن تا حداقل یک جدول در صفحه ایجاد شود
        # ----------------------------------------------------

        wait.until(
            EC.presence_of_element_located(
                (
                    By.TAG_NAME,
                    "table"
                )
            )
        )

        # ----------------------------------------------------
        # کمی صبر برای کامل شدن اطلاعات JavaScript
        # ----------------------------------------------------

        time.sleep(3)

        # ----------------------------------------------------
        # پیدا کردن تمام جدول های صفحه
        # ----------------------------------------------------

        tables = driver.find_elements(
            By.TAG_NAME,
            "table"
        )

        print(
            "\nتعداد جدول های پیدا شده:",
            len(tables)
        )

        # ----------------------------------------------------
        # جستجوی مستقل هر صندوق
        # ----------------------------------------------------

        for target_symbol in TARGET_SYMBOLS:

            normalized_target = normalize_text(
                target_symbol
            )

            found = False

            # بررسی تمام جدول ها
            for table in tables:

                # گرفتن تمام ردیف های جدول
                rows = table.find_elements(
                    By.XPATH,
                    ".//tbody/tr"
                )

                # بررسی تمام ردیف ها
                for row in rows:

                    try:

                        # گرفتن تمام سلول های ردیف
                        cells = row.find_elements(
                            By.TAG_NAME,
                            "td"
                        )

                        # حداقل باید 14 ستون داشته باشیم
                        if len(cells) < 14:
                            continue

                        # ستون اول = نماد
                        symbol_text = cells[0].text.strip()

                        normalized_symbol = normalize_text(
                            symbol_text
                        )

                        # مقایسه نام استاندارد شده
                        if normalized_symbol == normalized_target:

                            # ستون 12 = NAV
                            nav = cells[11].text.strip()

                            # ستون 14 = درصد حباب
                            bubble = cells[13].text.strip()

                            tradersarena_results[target_symbol] = {

                                "symbol": symbol_text,

                                "bubble": bubble,

                                "nav": nav
                            }

                            found = True

                            print(
                                f"\n✅ {target_symbol} در TradersArena پیدا شد."
                            )

                            break

                    except Exception:

                        continue

                if found:
                    break

            # اگر پیدا نشد
            if not found:

                tradersarena_results[target_symbol] = None

                print(
                    f"\n❌ {target_symbol} در TradersArena پیدا نشد."
                )

        return tradersarena_results

    except Exception as e:

        print(
            "\n❌ خطا در دریافت اطلاعات TradersArena:"
        )

        print(e)

        return None

    finally:

        driver.quit()


# ============================================================
# مرحله 4
# نمایش خروجی نهایی
# ============================================================

def print_final_results(
    tsetmc_results,
    tradersarena_results
):

    print("\n")
    print("#" * 60)
    print("#                  خروجی نهایی")
    print("#" * 60)

    # --------------------------------------------------------
    # هر صندوق به صورت کامل و مستقل نمایش داده می شود
    # --------------------------------------------------------

    for target_symbol in TARGET_SYMBOLS:

        tsetmc = tsetmc_results.get(
            target_symbol
        )

        tradersarena = tradersarena_results.get(
            target_symbol
        )

        print("\n")
        print("=" * 60)

        print(
            f"صندوق: {target_symbol}"
        )

        print("=" * 60)

        # ----------------------------------------------------
        # اطلاعات TSETMC
        # ----------------------------------------------------

        if tsetmc is not None:

            print(
                f"نماد: {tsetmc['symbol']}"
            )

            print(
                f"نام: {tsetmc['name']}"
            )

            print(
                f"تعداد معاملات: {tsetmc['transactions']}"
            )

            print(
                f"حجم: {tsetmc['volume']}"
            )

            print(
                f"ارزش: {tsetmc['value']}"
            )

            print(
                f"دیروز: {tsetmc['yesterday']}"
            )

            print(
                f"اولین: {tsetmc['first']}"
            )

            print(
                f"آخرین معامله: {tsetmc['last']}"
            )

            print(
                f"تغییر آخرین معامله: {tsetmc['last_change']}"
            )

            print(
                f"درصد آخرین معامله: {tsetmc['last_percent']}"
            )

            print(
                f"قیمت پایانی: {tsetmc['closing']}"
            )

            print(
                f"تغییر قیمت پایانی: {tsetmc['closing_change']}"
            )

            print(
                f"درصد قیمت پایانی: {tsetmc['closing_percent']}"
            )

        else:

            print(
                "❌ اطلاعات TSETMC پیدا نشد."
            )

        # ----------------------------------------------------
        # اطلاعات TradersArena
        # ----------------------------------------------------

        if tradersarena is not None:

            print(
                f"درصد حباب: {tradersarena['bubble']}"
            )

            print(
                f"NAV: {tradersarena['nav']}"
            )

        else:

            print(
                "❌ اطلاعات TradersArena پیدا نشد."
            )


# ============================================================
# اجرای اصلی برنامه
# ============================================================

if __name__ == "__main__":

    # --------------------------------------------------------
    # مرحله 1 - دانلود
    # --------------------------------------------------------

    downloaded_file = download_market_watch()

    if downloaded_file is None:

        print(
            "\n❌ برنامه به دلیل خطا در دانلود متوقف شد."
        )

    else:

        # ----------------------------------------------------
        # مرحله 2 - TSETMC
        # ----------------------------------------------------

        tsetmc_results = find_all_symbols(
            downloaded_file
        )

        if tsetmc_results is None:

            print(
                "\n❌ برنامه به دلیل خطا در خواندن Excel متوقف شد."
            )

        else:

            # ------------------------------------------------
            # مرحله 3 - TradersArena
            # ------------------------------------------------

            tradersarena_results = get_tradersarena_data()

            if tradersarena_results is None:

                print(
                    "\n❌ اطلاعات TradersArena دریافت نشد."
                )

                # ایجاد خروجی خالی برای هر سه صندوق
                tradersarena_results = {
                    symbol: None
                    for symbol in TARGET_SYMBOLS
                }

            # ------------------------------------------------
            # مرحله 4 - خروجی نهایی
            # ------------------------------------------------

            print_final_results(
                tsetmc_results,
                tradersarena_results
            )

    print("\n")
    print("#" * 60)
    print("#                  پایان برنامه")
    print("#" * 60)