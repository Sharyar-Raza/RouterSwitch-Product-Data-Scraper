import os
import re
import json
import time
import random
import ctypes
import requests
import pandas as pd
import winreg
import undetected_chromedriver as uc

from io import BytesIO
from PIL import Image
from urllib.parse import urljoin, quote

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# WINDOWS SLEEP PREVENTION
# ============================================================

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001


def prevent_windows_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED
        )
        print("Windows sleep prevention: ON")
    except Exception as e:
        print("Sleep prevention error:", e)


def restore_windows_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        print("Windows sleep prevention: OFF")
    except Exception:
        pass


# ============================================================
# WEBSITE
# ============================================================

BASE_URL = "https://www.router-switch.com/"
SEARCH_URL = "https://www.router-switch.com/search/"


# ============================================================
# SAME FOLDER EXCEL + OUTPUT
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_CANDIDATES = [
    
    "ORM_Cisco_Switches_Full_Categorization_1_cleaned.xlsx",
]

TARGET_SHEET = "Master (4000+ Switches)"
TARGET_SKU_COLUMN = "SKU / Part Number"


def find_master_excel():
    for filename in EXCEL_CANDIDATES:
        path = os.path.join(SCRIPT_DIR, filename)

        if os.path.exists(path):
            return path

    for filename in os.listdir(SCRIPT_DIR):
        lower = filename.lower()

        if not lower.endswith(".xlsx"):
            continue

        if lower.startswith("~$"):
            continue

        if "result" in lower:
            continue

        return os.path.join(
            SCRIPT_DIR,
            filename
        )

    raise FileNotFoundError(
        "Master Excel not found in same folder as scrapt.py"
    )


MASTER_EXCEL = find_master_excel()

PROJECT_ROOT = SCRIPT_DIR

IMAGE_ROOT = os.path.join(
    PROJECT_ROOT,
    "images"
)

OUTPUT_EXCEL = os.path.join(
    PROJECT_ROOT,
    "routerswitch_result.xlsx"
)

CHECKPOINT_FILE = os.path.join(
    PROJECT_ROOT,
    "routerswitch_checkpoint.txt"
)

SUCCESS_LOG = os.path.join(
    PROJECT_ROOT,
    "routerswitch_success.txt"
)

FAILED_LOG = os.path.join(
    PROJECT_ROOT,
    "routerswitch_failed.txt"
)

FULL_LOG = os.path.join(
    PROJECT_ROOT,
    "routerswitch_full_log.txt"
)

PROFILE_DIR = os.path.join(
    PROJECT_ROOT,
    "chrome_profile_routerswitch"
)

os.makedirs(
    IMAGE_ROOT,
    exist_ok=True
)

os.makedirs(
    PROFILE_DIR,
    exist_ok=True
)


# ============================================================
# SETTINGS
# ============================================================

PAGE_TIMEOUT = 45
WAIT_TIME = 12
PAGE_RETRIES = 3
RETRY_DELAY = 5

SEARCH_WAIT = 1.5
PRODUCT_WAIT = 1.0

# Cloudflare par zyada der stuck nahi rahega
CLOUDFLARE_MAX_WAIT = 240
CLOUDFLARE_CHECK_INTERVAL = 2

IMAGE_TIMEOUT = 25

MIN_WIDTH = 100
MIN_HEIGHT = 100
MIN_BYTES = 2000

SLEEP_BETWEEN_SKUS_MIN = 2.0
SLEEP_BETWEEN_SKUS_MAX = 4.0

LAST_PAGE_ERROR = ""

FINAL_STATUSES = [
    "DONE",
    "NOT FOUND",
]


# ============================================================
# HELPERS
# ============================================================

def log(message, level="INFO"):

    line = f"[{level}] {message}"

    print(line)

    try:
        with open(
            FULL_LOG,
            "a",
            encoding="utf-8"
        ) as f:

            f.write(
                line + "\n"
            )

    except Exception:
        pass

    target = None

    if level == "SUCCESS":
        target = SUCCESS_LOG

    elif level in {
        "NOT_FOUND",
        "NO_IMAGES",
        "ERROR",
        "RETRY"
    }:
        target = FAILED_LOG

    if target:

        try:
            with open(
                target,
                "a",
                encoding="utf-8"
            ) as f:

                f.write(
                    line + "\n"
                )

        except Exception:
            pass


def exact_sku(value):

    if value is None:
        return ""

    value = str(
        value
    ).strip()

    if value.lower() == "nan":
        return ""

    return value.upper()


def clean_filename(value):

    value = str(
        value
    ).strip()

    value = re.sub(
        r'[\\/:*?"<>|]',
        "_",
        value
    )

    value = re.sub(
        r"\s+",
        "_",
        value
    )

    return value


def normalize_url(
    url,
    base=BASE_URL
):

    if not url:
        return ""

    url = str(
        url
    ).strip()

    if url.startswith("//"):
        return "https:" + url

    return urljoin(
        base,
        url
    )


def compare_sku(
    requested_sku,
    website_sku
):

    requested = exact_sku(
        requested_sku
    )

    actual = exact_sku(
        website_sku
    )

    return bool(
        requested
        and actual
        and requested == actual
    )


# ============================================================
# CHECKPOINT
# ============================================================

def save_checkpoint(
    next_index
):

    temp = (
        CHECKPOINT_FILE
        + ".tmp"
    )

    with open(
        temp,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(
            str(
                next_index
            )
        )

    os.replace(
        temp,
        CHECKPOINT_FILE
    )


def load_checkpoint():

    if not os.path.exists(
        CHECKPOINT_FILE
    ):

        return 0

    try:

        with open(
            CHECKPOINT_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            value = f.read().strip()

        return (
            int(value)
            if value
            else 0
        )

    except Exception:
        return 0


# ============================================================
# CHROME
# ============================================================

def get_chrome_major():

    registry_paths = [

        (
            winreg.HKEY_CURRENT_USER,
            r"Software\Google\Chrome\BLBeacon"
        ),

        (
            winreg.HKEY_LOCAL_MACHINE,
            r"Software\Google\Chrome\BLBeacon"
        ),

        (
            winreg.HKEY_LOCAL_MACHINE,
            r"Software\WOW6432Node\Google\Chrome\BLBeacon"
        ),

    ]

    for hive, path in registry_paths:

        try:

            key = winreg.OpenKey(
                hive,
                path
            )

            version, _ = winreg.QueryValueEx(
                key,
                "version"
            )

            winreg.CloseKey(
                key
            )

            if version:

                return int(
                    str(
                        version
                    ).split(
                        "."
                    )[0]
                )

        except Exception:
            pass

    return None


def start_driver():

    options = uc.ChromeOptions()

    options.add_argument(
        "--start-maximized"
    )

    options.add_argument(
        f"--user-data-dir={PROFILE_DIR}"
    )

    options.add_argument(
        "--disable-notifications"
    )

    options.add_argument(
        "--disable-popup-blocking"
    )

    major = get_chrome_major()

    if major:

        driver = uc.Chrome(
            options=options,
            version_main=major,
            use_subprocess=True
        )

    else:

        driver = uc.Chrome(
            options=options,
            use_subprocess=True
        )

    driver.set_page_load_timeout(
        PAGE_TIMEOUT
    )

    return driver


def close_driver(
    driver
):

    if driver is None:
        return

    try:
        driver.quit()

    except Exception:
        pass


# ============================================================
# CLOUDFLARE
# ============================================================

def page_is_cloudflare(
    driver
):

    try:

        title = (
            driver.title
            or ""
        ).strip().lower()

        current_url = (
            driver.current_url
            or ""
        ).lower()

        if title in {
            "just a moment...",
            "just a moment"
        }:

            return True

        if (
            "/cdn-cgi/challenge-platform/"
            in current_url
        ):

            return True

        selectors = [

            "#challenge-running",

            "#challenge-stage",

            "#challenge-form",

            "iframe[src*='challenges.cloudflare.com']",

            "iframe[src*='turnstile']",

            ".cf-challenge",

        ]

        for selector in selectors:

            try:

                if driver.find_elements(
                    By.CSS_SELECTOR,
                    selector
                ):

                    return True

            except Exception:
                pass

        body = (
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
            or ""
        ).lower()

        markers = [

            "performing security verification",

            "verifying you are human",

            "checking your browser before accessing",

            "please wait while we verify",

        ]

        return any(
            marker in body[:5000]
            for marker in markers
        )

    except Exception:
        return False


def wait_for_cloudflare(
    driver
):

    global LAST_PAGE_ERROR

    if not page_is_cloudflare(
        driver
    ):

        return True

    print("")
    print(
        "Security verification detected."
    )

    print(
        f"Waiting maximum {CLOUDFLARE_MAX_WAIT} seconds..."
    )

    start = time.time()

    while (
        time.time()
        - start
        < CLOUDFLARE_MAX_WAIT
    ):

        if not page_is_cloudflare(
            driver
        ):

            print(
                "Verification cleared. Continuing..."
            )

            LAST_PAGE_ERROR = ""

            return True

        time.sleep(
            CLOUDFLARE_CHECK_INTERVAL
        )

    LAST_PAGE_ERROR = (
        f"Cloudflare verification timeout "
        f"after {CLOUDFLARE_MAX_WAIT}s"
    )

    print(
        "Security verification still active -> RETRY LATER"
    )

    return False


# ============================================================
# OPEN PAGE
# ============================================================

def open_page(
    driver,
    url,
    retries=PAGE_RETRIES
):

    global LAST_PAGE_ERROR

    LAST_PAGE_ERROR = ""

    for attempt in range(
        1,
        retries + 1
    ):

        try:

            print(
                f"Opening ({attempt}/{retries}):"
            )

            print(
                url
            )

            driver.get(
                url
            )

            WebDriverWait(
                driver,
                WAIT_TIME
            ).until(

                EC.presence_of_element_located(
                    (
                        By.TAG_NAME,
                        "body"
                    )
                )

            )

            if page_is_cloudflare(
                driver
            ):

                return wait_for_cloudflare(
                    driver
                )

            LAST_PAGE_ERROR = ""

            return True

        except Exception as e:

            LAST_PAGE_ERROR = (
                "Page/network/browser error: "
                + str(e)[:300]
            )

            print(
                "Page error:",
                str(e)[:300]
            )

            if attempt < retries:

                time.sleep(
                    RETRY_DELAY
                )

    return False


def search_page_has_no_results(
    driver
):

    try:

        body = (
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
            or ""
        ).lower()

        markers = [

            "no results found",

            "0 results for",

            "no products found",

            "your search returned no results",

        ]

        return any(
            marker in body
            for marker in markers
        )

    except Exception:
        return False


def page_is_404(
    driver
):

    try:

        title = (
            driver.title
            or ""
        ).lower()

        body = (
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
            or ""
        ).lower()

        markers = [

            "404",

            "page not found",

            "product not found",

        ]

        return any(
            marker in title
            or marker in body[:3000]
            for marker in markers
        )

    except Exception:
        return False


# ============================================================
# EXACT SEARCH RESULTS
# ============================================================

SEARCH_RESULT_XPATHS = [

    '//*[@id="maincontent"]//ol//li//h2/a',

    '//*[@id="maincontent"]//h2/a',

]


def collect_exact_search_candidates(
    driver,
    requested_sku
):

    requested_exact = exact_sku(
        requested_sku
    )

    candidates = {}

    for xpath in SEARCH_RESULT_XPATHS:

        try:

            links = driver.find_elements(
                By.XPATH,
                xpath
            )

        except Exception:
            continue

        for link in links:

            try:

                title = (

                    link.text

                    or link.get_attribute(
                        "textContent"
                    )

                    or ""

                ).strip()

                href = normalize_url(

                    link.get_attribute(
                        "href"
                    )

                )

                if not title:
                    continue

                if not href:
                    continue

                # Search title me comma se pehle SKU
                website_search_sku = (

                    title
                    .split(
                        ",",
                        1
                    )[0]
                    .strip()

                )

                print("")
                print(
                    "Search result title:"
                )
                print(
                    title
                )

                print(
                    "Excel SKU:",
                    requested_sku
                )

                print(
                    "Website SKU before comma:",
                    website_search_sku
                )

                if (
                    exact_sku(
                        website_search_sku
                    )
                    != requested_exact
                ):

                    print(
                        "Rejected: exact search SKU mismatch."
                    )

                    continue

                clean_href = (

                    href
                    .split("?")[0]
                    .split("#")[0]

                )

                candidates[
                    clean_href
                ] = (
                    website_search_sku,
                    title
                )

                print(
                    "EXACT SEARCH SKU MATCH"
                )

            except Exception:
                pass

    return [

        (
            href,
            website_sku,
            title
        )

        for href, (
            website_sku,
            title
        )

        in candidates.items()

    ]


# ============================================================
# PRODUCT PAGE SKU
# ============================================================

def extract_product_model(
    driver
):

    try:

        body = (
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
            or ""
        )

        match = re.search(

            r"\bModel\s*:\s*"
            r"([A-Za-z0-9._/+()-]+)",

            body,

            flags=re.I

        )

        if match:

            return (
                match.group(
                    1
                )
                .strip()
            )

    except Exception:
        pass


    xpaths = [

        "//th[translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz')='model']"
        "/following-sibling::td[1]",

        "//td[translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz')='model']"
        "/following-sibling::td[1]",

        "//dt[translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
        "'abcdefghijklmnopqrstuvwxyz')='model']"
        "/following-sibling::dd[1]",

    ]


    for xpath in xpaths:

        try:

            elements = driver.find_elements(
                By.XPATH,
                xpath
            )

            for element in elements:

                value = (

                    element.text

                    or element.get_attribute(
                        "textContent"
                    )

                    or ""

                ).strip()

                if value:

                    return (

                        value
                        .splitlines()[0]
                        .strip()

                    )

        except Exception:
            pass


    # JSON-LD fallback
    try:

        scripts = driver.find_elements(

            By.XPATH,

            "//script[@type='application/ld+json']"

        )

        for script in scripts:

            try:

                raw = (

                    script.get_attribute(
                        "innerHTML"
                    )

                    or ""

                ).strip()

                if not raw:
                    continue

                data = json.loads(
                    raw
                )

                queue = [
                    data
                ]

                while queue:

                    current = queue.pop()

                    if isinstance(
                        current,
                        dict
                    ):

                        obj_type = str(

                            current.get(
                                "@type",
                                ""
                            )

                        ).lower()

                        if (
                            "product"
                            in obj_type
                        ):

                            for key in [

                                "model",

                                "sku",

                                "mpn",

                                "productID",

                            ]:

                                value = current.get(
                                    key
                                )

                                if value:

                                    return str(
                                        value
                                    ).strip()

                        for value in current.values():

                            if isinstance(
                                value,
                                (
                                    dict,
                                    list
                                )
                            ):

                                queue.append(
                                    value
                                )

                    elif isinstance(
                        current,
                        list
                    ):

                        queue.extend(
                            current
                        )

            except Exception:
                pass

    except Exception:
        pass


    # H1 fallback
    try:

        h1s = driver.find_elements(
            By.XPATH,
            "//h1"
        )

        for h1 in h1s:

            text = (
                h1.text
                or ""
            ).strip()

            if text:

                return (

                    text
                    .split(
                        ",",
                        1
                    )[0]
                    .strip()

                )

    except Exception:
        pass

    return ""


# ============================================================
# SEARCH PRODUCT
# ============================================================

def search_product(
    driver,
    sku
):

    print("")
    print(
        "Searching Router-Switch SKU:"
    )

    print(
        sku
    )

    search_url = (
        SEARCH_URL
        + quote(
            str(
                sku
            ).strip(),
            safe=""
        )
    )

    print(
        "SEARCH URL:"
    )

    print(
        search_url
    )

    if not open_page(
        driver,
        search_url
    ):

        return (
            "RETRY",
            "",
            ""
        )

    time.sleep(
        SEARCH_WAIT
    )

    if search_page_has_no_results(
        driver
    ):

        return (
            "NOT_FOUND",
            "",
            ""
        )

    candidates = collect_exact_search_candidates(
        driver,
        sku
    )

    print(
        "Exact search candidates:",
        len(
            candidates
        )
    )

    if not candidates:

        return (
            "NOT_FOUND",
            "",
            ""
        )

    for (
        product_url,
        search_sku,
        title
    ) in candidates:

        print("")
        print(
            "Opening exact search match:"
        )

        print(
            product_url
        )

        if not open_page(
            driver,
            product_url
        ):

            return (
                "RETRY",
                "",
                ""
            )

        time.sleep(
            PRODUCT_WAIT
        )

        if page_is_404(
            driver
        ):

            continue

        page_sku = extract_product_model(
            driver
        )

        print("")
        print(
            "Excel SKU:",
            sku
        )

        print(
            "Search SKU:",
            search_sku
        )

        print(
            "Product page SKU:",
            page_sku
        )

        if not compare_sku(
            sku,
            page_sku
        ):

            print(
                "Rejected: product page Model mismatch."
            )

            continue

        print("")
        print(
            "========================================"
        )

        print(
            "EXACT SKU VERIFIED"
        )

        print(
            "========================================"
        )

        print(
            "SKU:",
            page_sku
        )

        print(
            "URL:",
            driver.current_url
        )

        return (
            "FOUND",
            driver.current_url,
            page_sku
        )

    return (
        "NOT_FOUND",
        "",
        ""
    )


# ============================================================
# REQUEST SESSION
# ============================================================

def create_session():

    session = requests.Session()

    session.headers.update(
        {

            "User-Agent":
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/140 Safari/537.36",

            "Accept":
                "image/avif,image/webp,image/apng,"
                "image/*,*/*;q=0.8",

        }
    )

    return session


def sync_cookies(
    driver,
    session
):

    try:

        session.cookies.clear()

        for cookie in driver.get_cookies():

            try:

                session.cookies.set(

                    cookie[
                        "name"
                    ],

                    cookie[
                        "value"
                    ]

                )

            except Exception:
                pass

        try:

            browser_ua = driver.execute_script(
                "return navigator.userAgent"
            )

            if browser_ua:

                session.headers.update(
                    {

                        "User-Agent":
                            str(
                                browser_ua
                            ).strip()

                    }
                )

        except Exception:
            pass

    except Exception:
        pass


# ============================================================
# IMAGE EXTRACTION HELPERS
# ============================================================

def image_urls_from_element(
    element,
    product_url
):

    urls = []

    attrs = [

        "currentSrc",

        "src",

        "data-src",

        "data-original",

        "data-lazy-src",

        "data-original-src",

        "data-image",

        "data-full",

        "data-full-image",

        "data-large",

        "data-large-image",

        "data-zoom-image",

        "data-src-big",

        "data-big",

        "data-zoom",

        "data-url",

        "data-origin",

    ]

    for attr in attrs:

        try:

            value = (

                element.get_attribute(
                    attr
                )

                or ""

            ).strip()

            if value:

                urls.append(

                    normalize_url(
                        value,
                        product_url
                    )

                )

        except Exception:
            pass

    for attr in [

        "srcset",

        "data-srcset"

    ]:

        try:

            value = (

                element.get_attribute(
                    attr
                )

                or ""

            ).strip()

            if not value:
                continue

            for item in value.split(
                ","
            ):

                bits = item.strip().split()

                if bits:

                    urls.append(

                        normalize_url(
                            bits[0],
                            product_url
                        )

                    )

        except Exception:
            pass

    return urls


def element_is_recommended_area(
    element
):

    try:

        ancestors = element.find_elements(

            By.XPATH,

            "./ancestor::*["
            "contains(translate(@class,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'recommend') "
            "or contains(translate(@id,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'recommend') "
            "or contains(translate(@class,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'related') "
            "or contains(translate(@id,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'related') "
            "or contains(translate(@class,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'similar') "
            "or contains(translate(@id,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'similar') "
            "or contains(translate(@class,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'upsell') "
            "or contains(translate(@id,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'upsell') "
            "or contains(translate(@class,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'cross-sell') "
            "or contains(translate(@id,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'cross-sell') "
            "or contains(translate(@class,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'recently') "
            "or contains(translate(@id,"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'recently')"
            "][1]"

        )

        return bool(
            ancestors
        )

    except Exception:
        return False


def probable_product_image(
    url
):

    if not url:
        return False

    lower = str(
        url
    ).lower()

    blocked = [

        "logo",

        "favicon",

        "sprite",

        "captcha",

        "payment",

        "paypal",

        "facebook",

        "instagram",

        "youtube",

        "twitter",

        "linkedin",

        "rating",

        "badge",

        "banner",

        "loader",

        "loading",

        "pixel",

        "avatar",

        "chat",

        "flag",

        "recommended",

        "recommendation",

        "related-product",

        "related_product",

        "similar-product",

        "similar_product",

        "upsell",

        "cross-sell",

        "cross_sell",

        "recently-viewed",

        "recently_viewed",

        "watermark",

        "placeholder",

        "no-image",

        "no_image",

        "noimage",

    ]

    if any(
        item in lower
        for item in blocked
    ):

        return False

    clean = (
        lower
        .split("?")[0]
        .split("#")[0]
    )

    if clean.endswith(
        (
            ".svg",
            ".gif"
        )
    ):

        return False

    return True


def extract_json_product_images(
    driver,
    product_url,
    requested_sku
):

    urls = []

    product_objects = []


    def walk(
        value
    ):

        if isinstance(
            value,
            dict
        ):

            obj_type = value.get(
                "@type",
                ""
            )

            if isinstance(
                obj_type,
                list
            ):

                object_types = [

                    str(
                        x
                    ).lower()

                    for x in obj_type

                ]

            else:

                object_types = [

                    str(
                        obj_type
                    ).lower()

                ]

            if any(

                x == "product"

                or x.endswith(
                    "/product"
                )

                for x in object_types

            ):

                product_objects.append(
                    value
                )

            for child in value.values():

                if isinstance(
                    child,
                    (
                        dict,
                        list
                    )
                ):

                    walk(
                        child
                    )

        elif isinstance(
            value,
            list
        ):

            for child in value:

                walk(
                    child
                )


    def add_image(
        value
    ):

        if isinstance(
            value,
            str
        ):

            normalized = normalize_url(
                value,
                product_url
            )

            if normalized:

                urls.append(
                    normalized
                )

        elif isinstance(
            value,
            list
        ):

            for item in value:

                add_image(
                    item
                )

        elif isinstance(
            value,
            dict
        ):

            for key in [

                "url",

                "contentUrl",

                "thumbnailUrl",

            ]:

                if key in value:

                    add_image(
                        value[
                            key
                        ]
                    )


    try:

        scripts = driver.find_elements(

            By.XPATH,

            "//script[@type='application/ld+json']"

        )

        for script in scripts:

            try:

                raw = (

                    script.get_attribute(
                        "innerHTML"
                    )

                    or ""

                ).strip()

                if raw:

                    walk(
                        json.loads(
                            raw
                        )
                    )

            except Exception:
                pass

    except Exception:
        pass


    exact_products = []

    for product in product_objects:

        identifiers = []

        for key in [

            "model",

            "sku",

            "mpn",

            "productID",

        ]:

            value = product.get(
                key
            )

            if value:

                identifiers.append(

                    str(
                        value
                    ).strip()

                )

        if any(

            compare_sku(
                requested_sku,
                identifier
            )

            for identifier
            in identifiers

        ):

            exact_products.append(
                product
            )


    if (
        not exact_products
        and len(
            product_objects
        ) == 1
    ):

        exact_products = [
            product_objects[
                0
            ]
        ]


    for product in exact_products:

        if "image" in product:

            add_image(
                product[
                    "image"
                ]
            )

        if "images" in product:

            add_image(
                product[
                    "images"
                ]
            )

    return urls


# ============================================================
# EXTRACT PRODUCT IMAGES
# ============================================================

def extract_product_images(
    driver,
    product_url,
    requested_sku
):

    current_sku = extract_product_model(
        driver
    )

    if not compare_sku(
        requested_sku,
        current_sku
    ):

        print(
            "BLOCKED IMAGE DOWNLOAD"
        )

        print(
            "Excel SKU:",
            requested_sku
        )

        print(
            "Website SKU:",
            current_sku
        )

        return []

    print("")
    print(
        "EXACT SKU VERIFIED FOR IMAGES:"
    )

    print(
        current_sku
    )

    print(
        "Exact product gallery only; related/recommended ignored."
    )

    urls = []


    try:

        for y in [

            0,

            250,

            500,

            750,

            1000

        ]:

            driver.execute_script(
                f"window.scrollTo(0,{y});"
            )

            time.sleep(
                0.2
            )

        driver.execute_script(
            "window.scrollTo(0,0);"
        )

        time.sleep(
            0.5
        )

    except Exception:
        pass


    def collect(
        xpath
    ):

        try:

            elements = driver.find_elements(
                By.XPATH,
                xpath
            )

            for element in elements:

                try:

                    if element_is_recommended_area(
                        element
                    ):

                        continue

                    urls.extend(

                        image_urls_from_element(
                            element,
                            product_url
                        )

                    )

                except Exception:
                    pass

        except Exception:
            pass


    selectors = [

        "//*[starts-with(@id,'magnifier-item-')]",

        "//*[starts-with(@id,'magnifier-item-')]//img",

        "//*[contains(@id,'magnifier')]//img",

        '//*[@id="maincontent"]'
        '//*[contains(@class,"gallery-placeholder")]//img',

        '//*[@id="maincontent"]'
        '//*[@data-gallery-role="gallery-placeholder"]//img',

        '//*[@id="maincontent"]'
        '//*[contains(@class,"fotorama")]//img',

        '//*[@id="maincontent"]'
        '//*[contains(@class,"fotorama__stage")]//img',

        '//*[@id="maincontent"]'
        '//*[contains(@class,"fotorama__nav")]//img',

        '//*[@id="maincontent"]'
        '//*[contains(@class,"product-image")]//img',

        '//*[@id="maincontent"]'
        '//*[contains(@class,"product-media")]//img',

    ]


    for xpath in selectors:

        collect(
            xpath
        )


    try:

        main_images = driver.find_elements(

            By.XPATH,

            '//*[@id="maincontent"]//img'

        )

        for image in main_images:

            try:

                if element_is_recommended_area(
                    image
                ):

                    continue

                descriptive = " ".join(
                    [

                        image.get_attribute(
                            "alt"
                        )
                        or "",

                        image.get_attribute(
                            "title"
                        )
                        or "",

                    ]
                )

                if (
                    requested_sku.lower()
                    in descriptive.lower()
                ):

                    urls.extend(

                        image_urls_from_element(
                            image,
                            product_url
                        )

                    )

            except Exception:
                pass

    except Exception:
        pass


    for selector in [

        "//meta[@property='og:image']",

        "//meta[@property='og:image:secure_url']",

        "//meta[@name='twitter:image']",

        "//link[@rel='image_src']",

    ]:

        try:

            elements = driver.find_elements(
                By.XPATH,
                selector
            )

            for element in elements:

                value = (

                    element.get_attribute(
                        "content"
                    )

                    or element.get_attribute(
                        "href"
                    )

                    or ""

                ).strip()

                if value:

                    urls.append(

                        normalize_url(
                            value,
                            product_url
                        )

                    )

        except Exception:
            pass


    urls.extend(

        extract_json_product_images(

            driver,

            product_url,

            requested_sku

        )

    )


    # ========================================================
    # IMPORTANT:
    # Duplicate image URLs intentionally preserved.
    # NO set()
    # NO hash duplicate check
    # ========================================================

    final_urls = []

    for image_url in urls:

        if not image_url:
            continue

        image_url = (

            str(
                image_url
            )

            .replace(
                "&amp;",
                "&"
            )

            .strip()

        )

        if not image_url.startswith(
            (
                "data:image/",
                "http://",
                "https://"
            )
        ):

            continue

        if not image_url.startswith(
            "data:image/"
        ):

            if not probable_product_image(
                image_url
            ):

                continue

        final_urls.append(
            image_url
        )

    print("")
    print(
        "EXACT PRODUCT IMAGE URLs "
        "(DUPLICATES ALLOWED):",
        len(
            final_urls
        )
    )

    return final_urls


# ============================================================
# DOWNLOAD IMAGE
# ============================================================

def download_image(
    session,
    image_url,
    product_url
):

    if str(
        image_url
    ).startswith(
        "data:image/"
    ):

        try:

            import base64

            header, encoded = str(
                image_url
            ).split(
                ",",
                1
            )

            if (
                ";base64"
                not in header.lower()
            ):

                return None

            raw = base64.b64decode(
                encoded
            )

            if len(
                raw
            ) < MIN_BYTES:

                return None

            image = Image.open(

                BytesIO(
                    raw
                )

            )

            image.load()

            if (
                image.width < MIN_WIDTH
                or image.height < MIN_HEIGHT
            ):

                return None

            return image

        except Exception:
            return None


    attempts = [

        {

            "Referer":
                product_url,

            "Accept":
                "image/avif,image/webp,"
                "image/apng,image/*,*/*;q=0.8",

        },

        {

            "Accept":
                "image/avif,image/webp,"
                "image/apng,image/*,*/*;q=0.8",

        },

    ]


    for attempt_number, headers in enumerate(
        attempts,
        start=1
    ):

        try:

            response = session.get(

                image_url,

                headers=headers,

                timeout=IMAGE_TIMEOUT,

                allow_redirects=True

            )

            print(

                f"Image HTTP attempt "
                f"{attempt_number}:",

                response.status_code

            )

            if (
                response.status_code
                != 200
            ):

                continue

            content_type = (

                response.headers.get(
                    "Content-Type",
                    ""
                )

                or ""

            ).lower()

            if (

                "text/html"
                in content_type

                or "application/json"
                in content_type

            ):

                continue

            if len(
                response.content
            ) < MIN_BYTES:

                continue

            image = Image.open(

                BytesIO(
                    response.content
                )

            )

            image.load()

            if (
                image.width < MIN_WIDTH
                or image.height < MIN_HEIGHT
            ):

                continue

            return image

        except Exception as e:

            print(
                "Image download error:",
                str(
                    e
                )[:200]
            )

    return None


# ============================================================
# SAVE PRODUCT IMAGES
# ============================================================

def save_product_images(
    driver,
    session,
    sku,
    product_url,
    image_urls
):

    """
    REQUIRED BEHAVIOR:

    - Old images NEVER delete.
    - Old images NEVER replace.
    - Duplicate images ARE ALLOWED.
    - Same image can save multiple times.
    - Every successful image gets next filename.
    """

    safe_sku = clean_filename(
        sku
    )

    sku_folder = os.path.join(
        IMAGE_ROOT,
        safe_sku
    )

    os.makedirs(
        sku_folder,
        exist_ok=True
    )

    print("")
    print(
        "SAVE FOLDER:"
    )

    print(
        sku_folder
    )

    print(
        "MODE: KEEP OLD IMAGES + DUPLICATES ALLOWED"
    )


    prefix = (
        f"{safe_sku}_RouterSwitch"
    ).lower()

    existing_paths = []

    used_numbers = set()


    try:

        for filename in sorted(
            os.listdir(
                sku_folder
            )
        ):

            filepath = os.path.join(
                sku_folder,
                filename
            )

            if (

                os.path.isfile(
                    filepath
                )

                and filename.lower().startswith(
                    prefix
                )

            ):

                existing_paths.append(
                    filepath
                )

                match = re.search(

                    r"_RouterSwitch_(\d+)\.jpg$",

                    filename,

                    flags=re.I

                )

                if match:

                    used_numbers.add(

                        int(
                            match.group(
                                1
                            )
                        )

                    )

    except Exception:
        pass


    thumbnail_path = os.path.join(

        sku_folder,

        f"{safe_sku}"
        f"_RouterSwitch_thumbnail.jpg"

    )

    next_number = 1

    new_paths = []

    new_source_urls = []


    for image_url in image_urls:

        print("")
        print(
            "Trying exact product image:"
        )

        print(
            str(
                image_url
            )[:1000]
        )

        image = download_image(

            session,

            image_url,

            product_url

        )

        if image is None:
            continue


        # First image only uses thumbnail if thumbnail does not exist.
        if not os.path.exists(
            thumbnail_path
        ):

            filepath = thumbnail_path

        else:

            while (
                next_number
                in used_numbers
            ):

                next_number += 1

            filepath = os.path.join(

                sku_folder,

                (
                    f"{safe_sku}"
                    f"_RouterSwitch_"
                    f"{next_number:02d}.jpg"
                )

            )

            used_numbers.add(
                next_number
            )

            next_number += 1


        # Never overwrite anything
        while os.path.exists(
            filepath
        ):

            while (
                next_number
                in used_numbers
            ):

                next_number += 1

            filepath = os.path.join(

                sku_folder,

                (
                    f"{safe_sku}"
                    f"_RouterSwitch_"
                    f"{next_number:02d}.jpg"
                )

            )

            used_numbers.add(
                next_number
            )

            next_number += 1


        try:

            image.convert(
                "RGB"
            ).save(

                filepath,

                "JPEG",

                quality=95

            )

        except Exception as e:

            print(
                "Image save error:",
                e
            )

            continue


        new_paths.append(
            filepath
        )

        new_source_urls.append(
            image_url
        )

        print(
            "IMAGE SAVED "
            "(DUPLICATES ALLOWED):"
        )

        print(
            filepath
        )


    final_paths = []

    try:

        for filename in sorted(
            os.listdir(
                sku_folder
            )
        ):

            filepath = os.path.join(
                sku_folder,
                filename
            )

            if (

                os.path.isfile(
                    filepath
                )

                and filename.lower().startswith(
                    prefix
                )

            ):

                final_paths.append(
                    filepath
                )

    except Exception:
        pass


    print("")
    print(
        "EXISTING IMAGES PRESERVED:",
        len(
            existing_paths
        )
    )

    print(
        "NEW IMAGES SAVED THIS RUN:",
        len(
            new_paths
        )
    )

    print(
        "TOTAL IMAGES IN SKU FOLDER:",
        len(
            final_paths
        )
    )


    return (
        final_paths,
        new_source_urls
    )


# ============================================================
# DATAFRAME
# ============================================================

def load_dataframe():

    if os.path.exists(
        OUTPUT_EXCEL
    ):

        try:

            return pd.read_excel(

                OUTPUT_EXCEL,

                sheet_name=TARGET_SHEET,

                dtype=str

            )

        except Exception:

            return pd.read_excel(

                OUTPUT_EXCEL,

                dtype=str

            )


    try:

        return pd.read_excel(

            MASTER_EXCEL,

            sheet_name=TARGET_SHEET,

            dtype=str

        )

    except Exception:

        return pd.read_excel(

            MASTER_EXCEL,

            dtype=str

        )


def prepare_columns(
    df
):

    columns = [

        "RouterSwitch Matched SKU",

        "RouterSwitch Product URL",

        "RouterSwitch Thumbnail Path",

        "RouterSwitch Additional Image Paths",

        "RouterSwitch Image Source URLs",

        "RouterSwitch Image Count",

        "RouterSwitch Summary",

        "RouterSwitch Status",

        "RouterSwitch Last Error",

    ]


    for column in columns:

        if column not in df.columns:

            df[
                column
            ] = ""


    df[
        "RouterSwitch Image Count"
    ] = pd.to_numeric(

        df[
            "RouterSwitch Image Count"
        ],

        errors="coerce"

    ).astype(
        "Int64"
    )


    return df


def save_result_excel(
    df
):

    temp = os.path.join(

        PROJECT_ROOT,

        "_routerswitch_temp.xlsx"

    )

    df.to_excel(

        temp,

        index=False,

        sheet_name=TARGET_SHEET

    )

    os.replace(

        temp,

        OUTPUT_EXCEL

    )


# ============================================================
# SUMMARY
# ============================================================

def print_final_summary(
    df
):

    status = (

        df[
            "RouterSwitch Status"
        ]

        .fillna(
            ""
        )

        .astype(
            str
        )

        .str.upper()

    )


    done = int(
        (
            status
            == "DONE"
        ).sum()
    )


    not_found = int(
        (
            status
            == "NOT FOUND"
        ).sum()
    )


    no_images = int(
        (
            status
            == "NO IMAGES"
        ).sum()
    )


    retry = int(
        (
            status
            == "RETRY"
        ).sum()
    )


    images = pd.to_numeric(

        df[
            "RouterSwitch Image Count"
        ],

        errors="coerce"

    ).fillna(
        0
    )


    print("")
    print(
        "================================================"
    )

    print(
        "ROUTER-SWITCH SUMMARY"
    )

    print(
        "================================================"
    )

    print(
        "TOTAL SKUs:",
        len(
            df
        )
    )

    print(
        "DONE:",
        done
    )

    print(
        "NO IMAGES:",
        no_images
    )

    print(
        "NOT FOUND:",
        not_found
    )

    print(
        "RETRY:",
        retry
    )

    print(
        "TOTAL IMAGE FILES:",
        int(
            images.sum()
        )
    )

    print(
        "OUTPUT EXCEL:"
    )

    print(
        OUTPUT_EXCEL
    )

    print(
        "IMAGE ROOT:"
    )

    print(
        IMAGE_ROOT
    )

    print(
        "================================================"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    prevent_windows_sleep()

    driver = None

    df = None

    try:

        df = load_dataframe()

        df = prepare_columns(
            df
        )


        if (
            TARGET_SKU_COLUMN
            not in df.columns
        ):

            raise RuntimeError(

                f"Excel column not found: "
                f"{TARGET_SKU_COLUMN}"

            )


        checkpoint = load_checkpoint()

        total = len(
            df
        )


        print("")
        print(
            "================================================"
        )

        print(
            "Excel:"
        )

        print(
            MASTER_EXCEL
        )

        print(
            "Sheet:",
            TARGET_SHEET
        )

        print(
            "SKU Column:",
            TARGET_SKU_COLUMN
        )

        print(
            "TOTAL SKUs:",
            total
        )

        print(
            "CHECKPOINT INDEX:",
            checkpoint
        )

        print(
            "REMAINING FROM CHECKPOINT:",
            max(
                total - checkpoint,
                0
            )
        )

        print(
            "NEXT / RESUME SKU:"
        )


        if checkpoint < total:

            print(

                exact_sku(

                    df.at[
                        checkpoint,
                        TARGET_SKU_COLUMN
                    ]

                )

            )

        else:

            print(
                "END"
            )


        print(
            "================================================"
        )


        session = create_session()

        driver = start_driver()


        for index in range(

            checkpoint,

            total

        ):


            sku = exact_sku(

                df.at[
                    index,
                    TARGET_SKU_COLUMN
                ]

            )


            if not sku:

                save_checkpoint(
                    index + 1
                )

                continue


            previous_status = str(

                df.at[
                    index,
                    "RouterSwitch Status"
                ]

            ).strip().upper()


            if (
                previous_status
                in FINAL_STATUSES
            ):

                print(

                    f"{index + 1}/{total} | "
                    f"{sku} | "
                    f"{previous_status} - SKIP"

                )

                save_checkpoint(
                    index + 1
                )

                continue


            print("")
            print(
                "#" * 48
            )

            print(

                f"SEARCHING SKU "
                f"{index + 1} "
                f"OF {total}"

            )

            print(
                f"Excel Row: {index + 2}"
            )

            print(
                "CURRENTLY SEARCHING SKU:"
            )

            print(
                sku
            )

            print(
                "REMAINING AFTER THIS:",
                total - (
                    index + 1
                )
            )

            print(
                "#" * 48
            )


            status, product_url, website_sku = search_product(

                driver,

                sku

            )


            # =================================================
            # RETRY / CLOUDFLARE
            # =================================================

            if status == "RETRY":

                df.at[
                    index,
                    "RouterSwitch Status"
                ] = "RETRY"

                df.at[
                    index,
                    "RouterSwitch Last Error"
                ] = LAST_PAGE_ERROR


                save_result_excel(
                    df
                )

                save_checkpoint(
                    index + 1
                )


                log(

                    f"SKU={sku} | "
                    f"RETRY | "
                    f"{LAST_PAGE_ERROR}",

                    "RETRY"

                )


                time.sleep(

                    random.uniform(

                        SLEEP_BETWEEN_SKUS_MIN,

                        SLEEP_BETWEEN_SKUS_MAX

                    )

                )

                continue


            # =================================================
            # NOT FOUND
            # =================================================

            if (
                status
                == "NOT_FOUND"
            ):

                df.at[
                    index,
                    "RouterSwitch Status"
                ] = "NOT FOUND"

                df.at[
                    index,
                    "RouterSwitch Last Error"
                ] = ""


                save_result_excel(
                    df
                )

                save_checkpoint(
                    index + 1
                )


                log(

                    f"SKU={sku} | "
                    f"NOT FOUND",

                    "NOT_FOUND"

                )


                time.sleep(

                    random.uniform(

                        SLEEP_BETWEEN_SKUS_MIN,

                        SLEEP_BETWEEN_SKUS_MAX

                    )

                )

                continue


            # =================================================
            # EXACT PRODUCT FOUND
            # =================================================

            df.at[
                index,
                "RouterSwitch Matched SKU"
            ] = website_sku


            df.at[
                index,
                "RouterSwitch Product URL"
            ] = product_url


            image_urls = extract_product_images(

                driver,

                product_url,

                sku

            )


            sync_cookies(

                driver,

                session

            )


            (
                paths,
                sources
            ) = save_product_images(

                driver,

                session,

                sku,

                product_url,

                image_urls

            )


            # =================================================
            # SUCCESS / IMAGES
            # =================================================

            if paths:

                df.at[
                    index,
                    "RouterSwitch Thumbnail Path"
                ] = paths[
                    0
                ]


                df.at[
                    index,
                    "RouterSwitch Additional Image Paths"
                ] = " | ".join(
                    paths[
                        1:
                    ]
                )


                df.at[
                    index,
                    "RouterSwitch Image Source URLs"
                ] = " | ".join(
                    sources
                )


                df.at[
                    index,
                    "RouterSwitch Image Count"
                ] = len(
                    paths
                )


                df.at[
                    index,
                    "RouterSwitch Summary"
                ] = (

                    f"{website_sku} | "
                    f"{product_url} | "
                    f"Images: {len(paths)}"

                )


                df.at[
                    index,
                    "RouterSwitch Status"
                ] = "DONE"


                df.at[
                    index,
                    "RouterSwitch Last Error"
                ] = ""


                log(

                    f"SKU={sku} | "
                    f"Images={len(paths)} | "
                    f"URL={product_url}",

                    "SUCCESS"

                )


            # =================================================
            # EXACT PRODUCT, NO IMAGE
            # =================================================

            else:

                df.at[
                    index,
                    "RouterSwitch Thumbnail Path"
                ] = ""


                df.at[
                    index,
                    "RouterSwitch Additional Image Paths"
                ] = ""


                df.at[
                    index,
                    "RouterSwitch Image Source URLs"
                ] = ""


                df.at[
                    index,
                    "RouterSwitch Image Count"
                ] = 0


                df.at[
                    index,
                    "RouterSwitch Summary"
                ] = (

                    f"{website_sku} | "
                    f"{product_url} | "
                    f"Images: 0"

                )


                df.at[
                    index,
                    "RouterSwitch Status"
                ] = "NO IMAGES"


                log(

                    f"SKU={sku} | "
                    f"EXACT FOUND BUT NO IMAGES | "
                    f"URL={product_url}",

                    "NO_IMAGES"

                )


            # =================================================
            # SAVE FIRST, CHECKPOINT SECOND
            # =================================================

            save_result_excel(
                df
            )

            save_checkpoint(
                index + 1
            )


            print("")
            print(
                "SKU COMPLETE"
            )

            print(
                "SKU:",
                website_sku
            )

            print(
                "URL:",
                product_url
            )

            print(
                "Images:",
                len(
                    paths
                )
            )

            print(
                "Next row:",
                index + 2
            )


            time.sleep(

                random.uniform(

                    SLEEP_BETWEEN_SKUS_MIN,

                    SLEEP_BETWEEN_SKUS_MAX

                )

            )


    except KeyboardInterrupt:

        print("")
        print(
            "SCRIPT STOPPED BY USER."
        )

        print(
            "Saved progress remains available."
        )


    except Exception as e:

        print("")
        print(
            "MAIN ERROR:"
        )

        print(
            e
        )

        if df is not None:

            try:

                save_result_excel(
                    df
                )

            except Exception:
                pass


    finally:

        close_driver(
            driver
        )

        restore_windows_sleep()


        if df is not None:

            try:

                save_result_excel(
                    df
                )

            except Exception as e:

                print(
                    "Final save warning:",
                    e
                )


            try:

                print_final_summary(
                    df
                )

            except Exception as e:

                print(
                    "Summary error:",
                    e
                )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    main()