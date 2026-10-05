import requests
from pathlib import Path
import gzip
import pandas as pd
import jdatetime
import time
import webbrowser
import json
import html

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
# تبدیل مقدار به عدد برای نمودار
# ============================================================

def to_number(value):

    try:

        if value is None:
            return 0

        text = str(value)

        # حذف جداکننده هزارگان
        text = text.replace(",", "")
        text = text.replace("٬", "")

        # حذف فاصله
        text = text.replace(" ", "")

        return float(text)

    except Exception:

        return 0


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

        # زمان دریافت موفق
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

    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

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

        # منتظر جدول
        wait.until(
            EC.presence_of_element_located(
                (
                    By.TAG_NAME,
                    "table"
                )
            )
        )

        # صبر برای تکمیل اطلاعات JavaScript
        time.sleep(3)

        # گرفتن تمام جدول ها
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

            for table in tables:

                rows = table.find_elements(
                    By.XPATH,
                    ".//tbody/tr"
                )

                for row in rows:

                    try:

                        cells = row.find_elements(
                            By.TAG_NAME,
                            "td"
                        )

                        if len(cells) < 14:
                            continue

                        # ستون اول = نماد
                        symbol_text = cells[0].text.strip()

                        normalized_symbol = normalize_text(
                            symbol_text
                        )

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
# ساخت داشبورد HTML
# ============================================================

def create_html_dashboard(
    tsetmc_results,
    tradersarena_results
):

    print("\n")
    print("=" * 60)
    print("مرحله 4 - ساخت داشبورد HTML")
    print("=" * 60)

    try:

        # ----------------------------------------------------
        # تاریخ و ساعت
        # ----------------------------------------------------

        now = jdatetime.datetime.now()

        jalali_date = now.strftime("%Y/%m/%d")
        current_time = now.strftime("%H:%M:%S")

        # ----------------------------------------------------
        # آماده سازی داده ها برای JavaScript
        # ----------------------------------------------------

        chart_closing = []
        chart_bubble = []

        for symbol in TARGET_SYMBOLS:

            tsetmc = tsetmc_results.get(symbol)
            traders = tradersarena_results.get(symbol)

            # قیمت پایانی
            if tsetmc is not None:

                chart_closing.append({
                    "name": symbol,
                    "value": to_number(
                        tsetmc["closing"]
                    )
                })

            # حباب
            if traders is not None:

                chart_bubble.append({
                    "name": symbol,
                    "value": to_number(
                        traders["bubble"]
                    )
                })

        closing_json = json.dumps(
            chart_closing,
            ensure_ascii=False
        )

        bubble_json = json.dumps(
            chart_bubble,
            ensure_ascii=False
        )

        # ----------------------------------------------------
        # ساخت کارت های صندوق
        # ----------------------------------------------------

        cards_html = ""

        for symbol in TARGET_SYMBOLS:

            tsetmc = tsetmc_results.get(symbol)
            traders = tradersarena_results.get(symbol)

            if tsetmc is None:

                cards_html += f"""
                <div class="fund-card">
                    <div class="fund-title">
                        {html.escape(symbol)}
                    </div>

                    <div class="not-found">
                        اطلاعات TSETMC پیدا نشد.
                    </div>
                </div>
                """

                continue

            # حباب
            if traders is not None:

                bubble = traders["bubble"]
                nav = traders["nav"]

                bubble_number = to_number(bubble)

                if bubble_number > 0:
                    bubble_class = "positive"
                elif bubble_number < 0:
                    bubble_class = "negative"
                else:
                    bubble_class = "neutral"

            else:

                bubble = "-"
                nav = "-"
                bubble_class = "neutral"

            cards_html += f"""
            <div class="fund-card">

                <div class="fund-title">
                    {html.escape(symbol)}
                </div>

                <div class="fund-name">
                    {html.escape(str(tsetmc["name"]))}
                </div>

                <div class="card-row">
                    <span>قیمت پایانی</span>
                    <strong>
                        {html.escape(str(tsetmc["closing"]))}
                    </strong>
                </div>

                <div class="card-row">
                    <span>آخرین معامله</span>
                    <strong>
                        {html.escape(str(tsetmc["last"]))}
                    </strong>
                </div>

                <div class="card-row">
                    <span>تغییر پایانی</span>
                    <strong>
                        {html.escape(str(tsetmc["closing_change"]))}
                    </strong>
                </div>

                <div class="card-row">
                    <span>درصد تغییر پایانی</span>
                    <strong>
                        {html.escape(str(tsetmc["closing_percent"]))}%
                    </strong>
                </div>

                <div class="card-row">
                    <span>درصد حباب</span>
                    <strong class="{bubble_class}">
                        {html.escape(str(bubble))}%
                    </strong>
                </div>

                <div class="card-row">
                    <span>NAV</span>
                    <strong>
                        {html.escape(str(nav))}
                    </strong>
                </div>

            </div>
            """

        # ----------------------------------------------------
        # ساخت جدول
        # ----------------------------------------------------

        table_rows = ""

        for symbol in TARGET_SYMBOLS:

            tsetmc = tsetmc_results.get(symbol)
            traders = tradersarena_results.get(symbol)

            if tsetmc is None:

                table_rows += f"""
                <tr>
                    <td>{html.escape(symbol)}</td>
                    <td colspan="14">
                        اطلاعات پیدا نشد
                    </td>
                </tr>
                """

                continue

            if traders is not None:

                bubble = traders["bubble"]
                nav = traders["nav"]

            else:

                bubble = "-"
                nav = "-"

            table_rows += f"""
            <tr>

                <td>
                    {html.escape(symbol)}
                </td>

                <td>
                    {html.escape(str(tsetmc["transactions"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["volume"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["value"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["yesterday"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["first"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["last"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["last_change"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["last_percent"]))}%
                </td>

                <td>
                    {html.escape(str(tsetmc["closing"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["closing_change"]))}
                </td>

                <td>
                    {html.escape(str(tsetmc["closing_percent"]))}%
                </td>

                <td>
                    {html.escape(str(bubble))}%
                </td>

                <td>
                    {html.escape(str(nav))}
                </td>

            </tr>
            """

        # ----------------------------------------------------
        # HTML کامل
        # ----------------------------------------------------

        html_content = f"""
<!DOCTYPE html>

<html lang="fa" dir="rtl">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>
داشبورد صندوق های دارویی
</title>

<style>

* {{
    box-sizing: border-box;
}}

body {{

    margin: 0;

    font-family:
        Tahoma,
        Arial,
        sans-serif;

    background:
        linear-gradient(
            135deg,
            #eef2f7,
            #f8fafc
        );

    color: #1e293b;
}}

.container {{

    width: 95%;

    max-width: 1500px;

    margin:
        0 auto;

    padding:
        30px 0 50px 0;
}}

.header {{

    background: white;

    border-radius: 18px;

    padding: 25px 30px;

    margin-bottom: 25px;

    box-shadow:
        0 8px 30px
        rgba(15, 23, 42, 0.08);

    border: 1px solid #e2e8f0;
}}

.header h1 {{

    margin: 0 0 10px 0;

    font-size: 28px;

    color: #0f172a;
}}

.header-info {{

    color: #64748b;

    font-size: 14px;
}}

.cards {{

    display: grid;

    grid-template-columns:
        repeat(
            3,
            minmax(0, 1fr)
        );

    gap: 20px;

    margin-bottom: 25px;
}}

.fund-card {{

    background: white;

    border-radius: 18px;

    padding: 22px;

    box-shadow:
        0 8px 30px
        rgba(15, 23, 42, 0.08);

    border:
        1px solid #e2e8f0;

    transition:
        transform 0.2s ease,
        box-shadow 0.2s ease;
}}

.fund-card:hover {{

    transform:
        translateY(-3px);

    box-shadow:
        0 14px 35px
        rgba(15, 23, 42, 0.12);
}}

.fund-title {{

    font-size: 24px;

    font-weight: bold;

    color: #0f172a;

    margin-bottom: 8px;
}}

.fund-name {{

    font-size: 13px;

    color: #64748b;

    margin-bottom: 18px;

    min-height: 20px;
}}

.card-row {{

    display: flex;

    justify-content: space-between;

    align-items: center;

    padding:
        9px 0;

    border-bottom:
        1px solid #f1f5f9;

    font-size: 14px;
}}

.card-row:last-child {{

    border-bottom: none;
}}

.card-row span {{

    color: #64748b;
}}

.card-row strong {{

    color: #0f172a;
}}

.positive {{

    color: #16a34a !important;
}}

.negative {{

    color: #dc2626 !important;
}}

.neutral {{

    color: #475569 !important;
}}

.section {{

    background: white;

    border-radius: 18px;

    padding: 25px;

    margin-bottom: 25px;

    box-shadow:
        0 8px 30px
        rgba(15, 23, 42, 0.08);

    border:
        1px solid #e2e8f0;
}}

.section-title {{

    font-size: 20px;

    font-weight: bold;

    margin-bottom: 20px;

    color: #0f172a;
}}

.table-wrapper {{

    overflow-x: auto;
}}

table {{

    width: 100%;

    border-collapse: collapse;

    min-width: 1200px;
}}

th {{

    background: #f1f5f9;

    color: #334155;

    font-size: 13px;

    padding: 13px 10px;

    border-bottom:
        2px solid #e2e8f0;

    white-space: nowrap;
}}

td {{

    padding: 13px 10px;

    text-align: center;

    border-bottom:
        1px solid #e2e8f0;

    font-size: 13px;

    white-space: nowrap;
}}

tr:hover td {{

    background: #f8fafc;
}}

.charts {{

    display: grid;

    grid-template-columns:
        repeat(
            2,
            minmax(0, 1fr)
        );

    gap: 20px;
}}

.chart-box {{

    background: #f8fafc;

    border-radius: 15px;

    padding: 20px;

    border:
        1px solid #e2e8f0;
}}

.chart-box h3 {{

    margin-top: 0;

    text-align: center;

    color: #334155;
}}

canvas {{

    width: 100% !important;

    height: 330px !important;
}}

.footer {{

    text-align: center;

    color: #94a3b8;

    font-size: 12px;

    margin-top: 25px;
}}

.not-found {{

    color: #dc2626;

    padding: 20px 0;
}}

@media (max-width: 1000px) {{

    .cards {{

        grid-template-columns:
            1fr;
    }}

    .charts {{

        grid-template-columns:
            1fr;
    }}

}}

</style>

<!-- Chart.js -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

</head>


<body>

<div class="container">


    <!-- ================================================ -->
    <!-- Header -->
    <!-- ================================================ -->

    <div class="header">

        <h1>
            داشبورد صندوق های دارویی
        </h1>

        <div class="header-info">

            آخرین بروزرسانی:

            {jalali_date}

            -

            {current_time}

            &nbsp;&nbsp; | &nbsp;&nbsp;

            منبع:

            TSETMC و TradersArena

        </div>

    </div>


    <!-- ================================================ -->
    <!-- Cards -->
    <!-- ================================================ -->

    <div class="cards">

        {cards_html}

    </div>


    <!-- ================================================ -->
    <!-- Table -->
    <!-- ================================================ -->

    <div class="section">

        <div class="section-title">

            جدول اطلاعات صندوق ها

        </div>

        <div class="table-wrapper">

            <table>

                <thead>

                    <tr>

                        <th>صندوق</th>

                        <th>تعداد معاملات</th>

                        <th>حجم</th>

                        <th>ارزش</th>

                        <th>دیروز</th>

                        <th>اولین</th>

                        <th>آخرین معامله</th>

                        <th>تغییر آخرین</th>

                        <th>درصد آخرین</th>

                        <th>قیمت پایانی</th>

                        <th>تغییر پایانی</th>

                        <th>درصد پایانی</th>

                        <th>درصد حباب</th>

                        <th>NAV</th>

                    </tr>

                </thead>

                <tbody>

                    {table_rows}

                </tbody>

            </table>

        </div>

    </div>


    <!-- ================================================ -->
    <!-- Charts -->
    <!-- ================================================ -->

    <div class="section">

        <div class="section-title">

            نمودارهای مقایسه ای

        </div>

        <div class="charts">


            <!-- نمودار قیمت پایانی -->

            <div class="chart-box">

                <h3>
                    مقایسه قیمت پایانی
                </h3>

                <canvas id="closingChart"></canvas>

            </div>


            <!-- نمودار حباب -->

            <div class="chart-box">

                <h3>
                    مقایسه درصد حباب
                </h3>

                <canvas id="bubbleChart"></canvas>

            </div>


        </div>

    </div>


    <div class="footer">

        داشبورد صندوق های دارویی

        <br>

        ساخته شده با Python

    </div>


</div>


<script>

const closingData =
    {closing_json};

const bubbleData =
    {bubble_json};


// ========================================================
// نمودار قیمت پایانی
// ========================================================

new Chart(

    document.getElementById(
        "closingChart"
    ),

    {{

        type: "bar",

        data: {{

            labels:
                closingData.map(
                    item => item.name
                ),

            datasets: [{{

                label:
                    "قیمت پایانی",

                data:
                    closingData.map(
                        item => item.value
                    ),

                borderWidth: 1

            }}]

        }},

        options: {{

            responsive: true,

            maintainAspectRatio: false,

            plugins: {{

                legend: {{

                    display: false

                }}

            }},

            scales: {{

                y: {{

                    beginAtZero: false

                }}

            }}

        }}

    }}

);


// ========================================================
// نمودار درصد حباب
// ========================================================

new Chart(

    document.getElementById(
        "bubbleChart"
    ),

    {{

        type: "bar",

        data: {{

            labels:
                bubbleData.map(
                    item => item.name
                ),

            datasets: [{{

                label:
                    "درصد حباب",

                data:
                    bubbleData.map(
                        item => item.value
                    ),

                borderWidth: 1

            }}]

        }},

        options: {{

            responsive: true,

            maintainAspectRatio: false,

            plugins: {{

                legend: {{

                    display: false

                }}

            }},

            scales: {{

                y: {{

                    title: {{

                        display: true,

                        text:
                            "درصد"

                    }}

                }}

            }}

        }}

    }}

);

</script>


</body>

</html>
"""

        # ----------------------------------------------------
        # نام فایل HTML
        # ----------------------------------------------------

        safe_date = now.strftime(
            "%Y-%m-%d"
        )

        safe_time = now.strftime(
            "%H-%M-%S"
        )

        html_file_name = (
            f"MedicineDashboard_"
            f"{safe_date}_"
            f"{safe_time}.html"
        )

        html_file_path = (
            SAVE_FOLDER /
            html_file_name
        )

        # ----------------------------------------------------
        # ذخیره HTML
        # ----------------------------------------------------

        with open(
            html_file_path,
            "w",
            encoding="utf-8"
        ) as file:

            file.write(html_content)

        print("\n" + "=" * 60)
        print("✅ داشبورد HTML با موفقیت ساخته شد")
        print("=" * 60)

        print("\nنام فایل:")
        print(html_file_name)

        print("\nمسیر:")
        print(html_file_path)

        print("=" * 60)

        # ----------------------------------------------------
        # باز کردن خودکار در مرورگر
        # ----------------------------------------------------

        print(
            "\nدر حال باز کردن داشبورد در مرورگر..."
        )

        webbrowser.open(
            html_file_path.resolve().as_uri()
        )

        print(
            "✅ داشبورد در مرورگر باز شد."
        )

        return html_file_path

    except Exception as e:

        print("\n")
        print("=" * 60)
        print("❌ خطا در ساخت داشبورد HTML")
        print("=" * 60)

        print("\nشرح خطا:")
        print(e)

        print("=" * 60)

        return None


# ============================================================
# اجرای اصلی برنامه
# ============================================================

if __name__ == "__main__":

    # --------------------------------------------------------
    # مرحله 1
    # --------------------------------------------------------

    downloaded_file = download_market_watch()

    if downloaded_file is None:

        print(
            "\n❌ برنامه به دلیل خطا در دانلود متوقف شد."
        )

    else:

        # ----------------------------------------------------
        # مرحله 2
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
            # مرحله 3
            # ------------------------------------------------

            tradersarena_results = get_tradersarena_data()

            if tradersarena_results is None:

                print(
                    "\n❌ اطلاعات TradersArena دریافت نشد."
                )

                tradersarena_results = {
                    symbol: None
                    for symbol in TARGET_SYMBOLS
                }

            # ------------------------------------------------
            # مرحله 4
            # ساخت داشبورد
            # ------------------------------------------------

            create_html_dashboard(
                tsetmc_results,
                tradersarena_results
            )

    print("\n")
    print("#" * 60)
    print("#                  پایان برنامه")
    print("#" * 60)