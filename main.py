#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AutoWeb-RPA-Demo
==================

一个面向学习的轻量级 RPA（Robotic Process Automation）演示项目。

它用 Selenium 驱动真实的 Chrome 浏览器，跑完一次完整的
"感知 -> 落盘 -> 操作 -> 校验" 闭环：

    第 1 步  感知：打开测试网页，定位并读取页面表格中的结构化数据；
    第 2 步  落盘：把数据整理成 DataFrame，同时导出 CSV 与格式化 Excel；
    第 3 步  操作：读取刚导出的数据，自动填入另一张表单并点击提交；
    第 4 步  校验：等待页面给出成功反馈，并截图留证。

RPA 的本质就是"用代码替代人的眼睛和手"：
    选择器  = 眼睛（定位元素）
    click / send_keys = 手（操作元素）
    断言 / 等待反馈  = 检查这件事到底做完没有

运行示例：
    python main.py                 # 跑公网公开测试站（需要能访问外网）
    python main.py --local         # 跑内置离线演示页（无需网络，适合录屏）
    python main.py --headless      # 无头模式，适合 CI
"""

from __future__ import annotations

import argparse
import logging
import sys
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# openpyxl 只用于 Excel 美化，属于"锦上添花"的依赖。
# 单独 try 一次，是为了在用户只装了 pandas 的情况下依然能产出 CSV，
# 而不是直接 ImportError 崩掉整个流程。
try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    HAS_OPENPYXL = True
except ImportError:  # pragma: no cover - 取决于运行环境
    HAS_OPENPYXL = False


# =============================================================================
# 一、常量与全局配置
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent      # 项目根目录，永远以本文件为锚点
DATA_DIR = BASE_DIR / "data"                    # 抓取结果与导出文件
SCREENSHOT_DIR = DATA_DIR / "screenshots"       # 出错现场 / 成功凭证
LOG_DIR = BASE_DIR / "logs"                     # 运行日志
DEMO_DIR = BASE_DIR / "demo_pages"              # 内置离线演示页（--local 模式用）

# ---- 公网目标站点：the-internet 是 Sauce Labs 维护多年的公开测试站，----
# ---- 页面静态、结构稳定，是练习 Selenium 的经典靶场。               ----
PUBLIC_TABLE_URL = "https://the-internet.herokuapp.com/tables"   # 一个含表格的静态页
PUBLIC_LOGIN_URL = "https://the-internet.herokuapp.com/login"    # 一个含表单的登录页

# the-internet 官方文档里公开的沙箱账号，仅用于练习，不涉及任何真实凭据
PUBLIC_USERNAME = "tomsmith"
PUBLIC_PASSWORD = "SuperSecretPassword!"
LOGIN_SUCCESS_MARK = "You logged into a secure area!"            # 登录成功的页面标志

# ---- 进程退出码：让 CI / 定时任务能区分"为什么失败"----
EXIT_OK = 0            # 全流程成功
EXIT_UNKNOWN = 1       # 未预料的异常
EXIT_TIMEOUT = 2       # 等待超时（页面慢 / 元素没出现）
EXIT_NOT_FOUND = 3     # 元素找不到（页面改版 / 选择器过期）
EXIT_BROWSER = 4       # 浏览器或驱动异常
EXIT_INTERRUPTED = 130  # 用户 Ctrl+C

logger = logging.getLogger("rpa")


# =============================================================================
# 二、自定义异常
# =============================================================================
class RpaDataError(RuntimeError):
    """页面能打开，但内容不符合预期（例如表格是空的、缺少表头）。

    这类错误和"元素找不到"不是一回事，单独定义便于上层给出更准确的提示。
    """


class RpaActionError(RuntimeError):
    """操作执行了，但页面反馈的结果不是成功态（例如登录被拒绝）。"""


# =============================================================================
# 三、运行配置
# =============================================================================
@dataclass
class RunConfig:
    """一次运行所需的全部参数，集中管理，避免到处散落魔法数字。"""

    mode: str                                  # 展示用的模式名，仅用于日志
    table_url: str                             # 数据来源页
    action_url: str                            # 要操作的录入页
    action_kind: str                           # "login"（公网登录）或 "registry"（离线登记表）
    headless: bool = False                     # 是否无头运行
    element_timeout: int = 10                  # 单个元素的最长等待（秒）
    page_timeout: int = 30                     # 页面加载的最长等待（秒）
    keep_open: bool = False                    # 跑完后是否保留浏览器窗口
    screenshot_on_success: bool = True         # 成功后是否截图留证
    extra: dict = field(default_factory=dict)  # 预留扩展位


# =============================================================================
# 四、日志
# =============================================================================
def setup_logging(verbose: bool = False) -> Path:
    """同时向控制台和文件输出日志，返回本次日志文件路径。

    写成函数而不是模块级 basicConfig，好处是：
    1. 日志文件可以按"每次运行"分文件，方便回溯历史；
    2. 避免被第三方库（如 urllib3）重复添加 handler 导致日志打印两遍。
    """
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"rpa_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()  # 幂等：重复调用不会叠加 handler

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s")
    )

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    console_handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")
    )

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return log_file


# =============================================================================
# 五、浏览器封装：把"眼睛"和"手"收进一个类里
# =============================================================================
class Browser:
    """对 Selenium WebDriver 的薄封装。

    为什么不直接在业务代码里调 driver？
    因为真实项目里 90% 的 Selenium 崩溃都来自两件事：
      a) 页面还没加载完就去点元素；
      b) 元素在可视区外 / 被固定导航栏遮挡。
    把显式等待和滚动兜底收敛到这里，业务代码就能保持干净。
    """

    def __init__(self, driver: webdriver.Chrome, element_timeout: int, page_timeout: int):
        self.driver = driver
        self.element_timeout = element_timeout
        self.page_timeout = page_timeout

    # ---------------------------- 页面级操作 ----------------------------
    def open(self, url: str) -> None:
        """打开页面，并额外兜底确认 DOM 已经就绪。

        driver.get() 本身会等 load 事件，但单页应用 / 慢站点经常在 load
        之后才渲染内容，所以这里再等一次 document.readyState。
        """
        logger.info("打开页面 -> %s", url)
        self.driver.get(url)
        WebDriverWait(self.driver, self.page_timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        logger.info("页面加载完成: %s", self.driver.title)

    # ---------------------------- 元素级操作 ----------------------------
    def find(self, by: str, value: str, timeout: int | None = None) -> WebElement:
        """等待元素"可见"后返回。

        用 visibility_of_element_located 而不是 presence_of_element_located，
        因为它能过滤掉那些"已经在 DOM 里、但还看不见"的元素，
        避免后续 click / send_keys 打在空气上。
        """
        wait = WebDriverWait(self.driver, timeout or self.element_timeout)
        try:
            return wait.until(EC.visibility_of_element_located((by, value)))
        except TimeoutException:
            logger.error("元素定位超时: by=%s, value=%s", by, value)
            raise  # 交给上层统一分类处理

    def type(self, by: str, value: str, text: str) -> None:
        """清空输入框并键入内容，模拟人工输入。"""
        element = self.find(by, value)
        element.clear()
        element.send_keys(text)
        logger.info("输入 -> %s = %r", value, text)

    def click(self, by: str, value: str) -> None:
        """滚动到可视区再点击。

        scrollIntoView 是为了规避 ElementClickInterceptedException：
        目标元素被吸顶导航、Cookie 弹窗挡住时，原生 click 会失败。
        """
        element = self.find(by, value)
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", element
        )
        element.click()
        logger.info("点击 -> %s", value)

    def wait_text(self, by: str, value: str, keyword: str, timeout: int | None = None) -> str:
        """等待某个元素出现，且其文本包含指定关键字，返回该文本。

        提交表单后页面往往需要一次请求才回写结果，
        用"等文本出现"代替"点完就取"，可以彻底消除竞态。
        """
        wait = WebDriverWait(self.driver, timeout or self.element_timeout)

        def _has_keyword(d: webdriver.Chrome) -> str | bool:
            try:
                element = d.find_element(by, value)
            except NoSuchElementException:
                return False
            text = element.text.strip()
            return text if keyword in text else False

        return wait.until(_has_keyword)

    def take_screenshot(self, tag: str) -> Path | None:
        """截图留证。失败也不抛异常，避免"截图挂了"掩盖真正的业务错误。"""
        try:
            SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
            path = SCREENSHOT_DIR / f"{tag}_{datetime.now():%Y%m%d_%H%M%S}.png"
            self.driver.save_screenshot(str(path))
            logger.info("截图已保存: %s", path)
            return path
        except Exception as exc:  # noqa: BLE001 - 截图失败不应影响主流程
            logger.warning("截图失败（不影响主流程）: %s", exc)
            return None


def build_driver(headless: bool) -> webdriver.Chrome:
    """创建 Chrome 实例。

    Selenium 4.6+ 内置 Selenium Manager，会自动匹配并下载 chromedriver，
    因此 requirements.txt 里不需要再写 webdriver-manager。
    """
    options = Options()
    if headless:
        options.add_argument("--headless=new")        # 新版无头模式，渲染更接近真实浏览器
    options.add_argument("--window-size=1440,900")     # 固定窗口，截图尺寸稳定
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")               # 兼容部分容器 / CI 环境
    options.add_argument("--disable-dev-shm-usage")    # 避免 /dev/shm 过小导致崩溃
    options.add_argument("--lang=zh-CN")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])  # 去掉控制台噪音

    logger.info("启动 Chrome（headless=%s）", headless)
    try:
        driver = webdriver.Chrome(options=options)
    except WebDriverException as exc:
        logger.error(
            "Chrome 启动失败：请确认已安装 Chrome 浏览器，且版本与 chromedriver 兼容。原始错误: %s",
            exc,
        )
        raise
    driver.set_page_load_timeout(60)  # 页面加载兜底超时，防止卡死在慢资源上
    return driver


# =============================================================================
# 六、抓取：把网页表格变成 DataFrame
# =============================================================================
def scrape_table(browser: Browser, table_id: str = "table1") -> pd.DataFrame:
    """读取指定表格，返回结构化 DataFrame。

    实现要点：
    1. 表头从 thead 拿，而不是硬编码列名 —— 页面加列了脚本也不会错位；
    2. 逐行按 td 收集文本，兼容 td 里嵌套 <a> 的情况（.text 会取到叶子文本）；
    3. 丢弃没有意义的 Action 列（里面只有 edit / delete 链接文字）；
    4. 追加 row_no / scraped_at 两列，方便追溯数据来源。
    """
    logger.info("开始抓取表格: #%s", table_id)
    table = browser.find(By.ID, table_id)

    headers = [th.text.strip() for th in table.find_elements(By.CSS_SELECTOR, "thead th")]
    if not headers:
        raise RpaDataError(f"表格 #{table_id} 没有 thead 表头，无法确定列名")

    raw_rows: list[list[str]] = []
    for tr in table.find_elements(By.CSS_SELECTOR, "tbody tr"):
        cells = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, "td")]
        if any(cells):  # 跳过纯空白行 / 分隔行
            raw_rows.append(cells)

    if not raw_rows:
        raise RpaDataError(f"表格 #{table_id} 中没有任何数据行")

    # 表头与单元格数量可能不一致（页面有合并单元格等），取两者最小值对齐，
    # 宁可少一列，也不要让 Pandas 因为长度不匹配而抛错。
    width = min(len(headers), min(len(row) for row in raw_rows))
    df = pd.DataFrame([row[:width] for row in raw_rows], columns=headers[:width])

    # 丢掉无意义的操作列
    dropped = [c for c in df.columns if c.lower() in {"action", "操作"}]
    if dropped:
        df = df.drop(columns=dropped)
        logger.info("已丢弃无意义列: %s", ", ".join(dropped))

    # 追加元数据列，让导出文件本身具备"可追溯性"
    df.insert(0, "row_no", range(1, len(df) + 1))
    df["scraped_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    logger.info("抓取完成：%d 行 x %d 列", df.shape[0], df.shape[1])
    return df


# =============================================================================
# 七、落盘：CSV + 格式化 Excel
# =============================================================================
def export_dataframe(df: pd.DataFrame, stem: str = "web_table_data") -> tuple[Path, Path]:
    """把 DataFrame 落盘为 CSV 与 XLSX，返回两个路径。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    csv_path = DATA_DIR / f"{stem}.csv"
    xlsx_path = DATA_DIR / f"{stem}.xlsx"

    # encoding="utf-8-sig"：带 BOM，双击用 Excel 打开中文不乱码。
    # 这是 Excel + 中文场景下最常见的坑，面试时可以直接讲。
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    logger.info("CSV 已生成: %s（%d 行）", csv_path, len(df))

    # 写 Excel 需要 openpyxl 引擎
    df.to_excel(xlsx_path, index=False, sheet_name="web_data")
    _beautify_workbook(xlsx_path)
    logger.info("Excel 已生成: %s（%d 行）", xlsx_path, len(df))
    return csv_path, xlsx_path


def _beautify_workbook(path: Path) -> None:
    """用 openpyxl 给导出的表加一层"能直接交付给业务方"的样式。

    内容包括：冻结首行、表头深蓝底白字、按内容自适应列宽、开启筛选。
    这些都是"脚本产物"和"能拿去汇报的产物"之间的差距。
    """
    if not HAS_OPENPYXL:
        logger.warning("未检测到 openpyxl，跳过 Excel 美化（CSV 不受影响）")
        return

    workbook = load_workbook(path)
    sheet = workbook.active

    sheet.freeze_panes = "A2"  # 冻结表头
    header_fill = PatternFill("solid", fgColor="1F4E79")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 简单估算列宽：取该列最长文本长度 + 余量，并限制上下界。
    # 注意中文是宽字符，这里用 len() 会低估，所以额外留了 +4 的余量。
    for index, column in enumerate(sheet.iter_cols(min_row=1, max_row=sheet.max_row), start=1):
        longest = max((len(str(c.value)) for c in column if c.value is not None), default=8)
        sheet.column_dimensions[get_column_letter(index)].width = min(max(longest + 4, 12), 42)

    sheet.auto_filter.ref = sheet.dimensions  # 表头筛选，业务方最爱
    workbook.save(path)


# =============================================================================
# 八、操作：自动录入表单
# =============================================================================
def login_action(browser: Browser) -> str:
    """场景 A：在公开测试站完成一次模拟登录。

    对应真实工作中的"登录后台系统"环节。失败时页面会渲染 .flash.error，
    这里先取 flash 文案再判断，比直接等 .success 能得到更清晰的错误原因。
    """
    browser.type(By.ID, "username", PUBLIC_USERNAME)
    browser.type(By.ID, "password", PUBLIC_PASSWORD)
    browser.click(By.CSS_SELECTOR, "button[type='submit']")

    # 登录后页面会跳转到 /secure，并渲染一条 #flash 提示
    message = browser.find(By.ID, "flash").text.strip()
    if LOGIN_SUCCESS_MARK not in message:
        raise RpaActionError(f"登录未成功，页面提示：{message}")
    logger.info("登录成功：%s", message.replace("\n", " "))
    return message


def registry_action(browser: Browser, df: pd.DataFrame) -> str:
    """场景 B：读取刚抓取的数据，回填到一张登记表单并提交。

    这才是"数据搬运"最完整的形态：源数据 -> 本地文件 -> 目标表单。
    这里取表格中的第一行作为示例录入对象（真实业务里应遍历全部行）。
    """
    if df.empty:
        raise RpaDataError("没有可录入的数据：DataFrame 为空")

    # 从抓取结果里取值。用 .get() 兜底，避免源表少列时直接 KeyError。
    record = df.iloc[0].to_dict()
    first_name = str(record.get("First Name", "")).strip()
    last_name = str(record.get("Last Name", "")).strip()
    full_name = f"{first_name} {last_name}".strip() or "未命名联系人"
    email = str(record.get("Email", "")).strip()

    logger.info("准备录入第 1 条数据：%s / %s", full_name, email)
    browser.type(By.ID, "cust_name", full_name)
    browser.type(By.ID, "cust_email", email)
    browser.type(
        By.ID,
        "note",
        f"由 AutoWeb-RPA-Demo 自动录入；本次共抓取 {len(df)} 条源数据。",
    )
    browser.click(By.ID, "submit-btn")

    message = browser.wait_text(By.ID, "result", "提交成功")
    logger.info("表单提交成功：%s", message)
    return message


# =============================================================================
# 九、主流程
# =============================================================================
def run(config: RunConfig) -> int:
    """执行一次完整 RPA 流程，返回进程退出码。

    结构上刻意把 try / except / finally 分层写清楚：
      - finally 负责"无论如何都要关浏览器"，防止进程残留；
      - except 按异常类型分流，让日志和退出码都能定位到责任方。
    """
    driver: webdriver.Chrome | None = None
    browser: Browser | None = None

    try:
        logger.info("=" * 64)
        logger.info("AutoWeb-RPA-Demo 启动 | 模式=%s | headless=%s", config.mode, config.headless)
        logger.info("=" * 64)

        # ---------------- 第 0 步：启动浏览器 ----------------
        driver = build_driver(config.headless)
        browser = Browser(driver, config.element_timeout, config.page_timeout)

        # ---------------- 第 1 步：感知，抓取表格 ----------------
        logger.info("[1/4] 打开数据源页面并抓取表格")
        browser.open(config.table_url)
        df = scrape_table(browser, "table1")

        # ---------------- 第 2 步：落盘，生成 CSV / Excel ----------------
        logger.info("[2/4] 导出数据到本地文件")
        csv_path, xlsx_path = export_dataframe(df)

        # 回读校验：确认落盘文件的行数与内存中一致，避免"写了一半"的静默失败
        roundtrip = pd.read_csv(csv_path)
        if len(roundtrip) != len(df):
            raise RpaDataError(
                f"落盘校验失败：内存 {len(df)} 行，CSV 回读 {len(roundtrip)} 行"
            )
        logger.info("落盘校验通过：CSV 回读 %d 行，与内存一致", len(roundtrip))

        # ---------------- 第 3 步：操作，自动填入并提交 ----------------
        logger.info("[3/4] 打开录入页面并自动填写表单")
        browser.open(config.action_url)
        if config.action_kind == "login":
            result_message = login_action(browser)
        elif config.action_kind == "registry":
            result_message = registry_action(browser, df)
        else:
            raise RpaActionError(f"未知的录入类型: {config.action_kind}")

        # ---------------- 第 4 步：校验，截图留证 ----------------
        logger.info("[4/4] 校验页面反馈")
        if config.screenshot_on_success:
            browser.take_screenshot("success")

        logger.info("-" * 64)
        logger.info("RPA 流程执行成功")
        logger.info("  抓取数据: %d 行 -> %s", len(df), csv_path.name)
        logger.info("  Excel 文件: %s", xlsx_path.name)
        logger.info("  页面反馈: %s", result_message.replace("\n", " "))
        logger.info("-" * 64)
        return EXIT_OK

    # ---------------- 异常分流 ----------------
    # 注意顺序：TimeoutException / NoSuchElementException 都是 WebDriverException
    # 的子类，必须先捕获子类，否则会被父类一并吃掉。
    except TimeoutException as exc:
        logger.error("[超时] 等待页面或元素超时，可能是网络慢、页面改版或选择器过期")
        logger.debug("详细堆栈:\n%s", traceback.format_exc())
        if driver is not None:
            _safe_screenshot(driver, "error_timeout")
        logger.error("原始信息: %s", exc)
        return EXIT_TIMEOUT

    except NoSuchElementException as exc:
        logger.error("[元素缺失] 页面上找不到目标元素，请检查选择器是否仍然有效")
        logger.debug("详细堆栈:\n%s", traceback.format_exc())
        if driver is not None:
            _safe_screenshot(driver, "error_not_found")
        logger.error("原始信息: %s", exc)
        return EXIT_NOT_FOUND

    except WebDriverException as exc:
        logger.error("[浏览器异常] 浏览器 / 驱动层面出错，例如 Chrome 未安装、版本不匹配、页面崩溃")
        logger.debug("详细堆栈:\n%s", traceback.format_exc())
        logger.error("原始信息: %s", exc)
        return EXIT_BROWSER

    except (RpaDataError, RpaActionError) as exc:
        logger.error("[业务异常] %s", exc)
        if driver is not None:
            _safe_screenshot(driver, "error_business")
        return EXIT_UNKNOWN

    except KeyboardInterrupt:
        logger.warning("[中断] 收到 Ctrl+C，正在关闭浏览器")
        return EXIT_INTERRUPTED

    except Exception as exc:  # noqa: BLE001 - 兜底，保证异常一定有出处
        logger.error("[未知异常] %s", exc)
        logger.debug("详细堆栈:\n%s", traceback.format_exc())
        if driver is not None:
            _safe_screenshot(driver, "error_unexpected")
        return EXIT_UNKNOWN

    # ---------------- 收尾 ----------------
    finally:
        if driver is not None:
            if config.keep_open:
                try:
                    input("浏览器已保留，按回车键关闭并退出...")
                except (EOFError, KeyboardInterrupt):
                    pass  # 非交互环境（CI）里没有 stdin，直接放行
            try:
                driver.quit()
                logger.info("浏览器已关闭")
            except Exception:  # noqa: BLE001 - quit 失败不应改变最终退出码
                logger.debug("关闭浏览器时出现异常（可忽略）")


def _safe_screenshot(driver: webdriver.Chrome, tag: str) -> None:
    """异常路径里的截图兜底。

    这里再包一层 try，是因为出错时页面可能已经崩溃，
    save_screenshot 本身也可能抛异常，绝不能让它覆盖掉原始错误。
    """
    try:
        SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
        path = SCREENSHOT_DIR / f"{tag}_{datetime.now():%Y%m%d_%H%M%S}.png"
        driver.save_screenshot(str(path))
        logger.info("出错现场已截图: %s", path)
    except Exception:  # noqa: BLE001
        logger.debug("出错现场截图失败（忽略）")


# =============================================================================
# 十、命令入口
# =============================================================================
def build_run_config(args: argparse.Namespace) -> RunConfig:
    """根据命令行参数组装运行配置。"""
    common = dict(
        headless=args.headless,
        element_timeout=args.timeout,
        page_timeout=args.page_load_timeout,
        keep_open=args.keep_open,
    )

    if args.local:
        # 内置离线演示页：把 HTML 转成 file:// URI，不需要任何网络
        return RunConfig(
            mode="离线演示（demo_pages）",
            table_url=(DEMO_DIR / "table.html").as_uri(),
            action_url=(DEMO_DIR / "form.html").as_uri(),
            action_kind="registry",
            **common,
        )

    return RunConfig(
        mode="公网测试站（the-internet）",
        table_url=PUBLIC_TABLE_URL,
        action_url=PUBLIC_LOGIN_URL,
        action_kind="login",
        **common,
    )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="AutoWeb-RPA-Demo",
        description="轻量级 RPA 演示：网页表格抓取 + 自动填表提交",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  python main.py --local --headless   # CI / 录屏推荐，无需外网\n"
            "  python main.py --keep-open          # 跑完保留窗口，方便讲解\n"
            "  python main.py -v                   # 输出 DEBUG 级日志（含堆栈）"
        ),
    )
    parser.add_argument("--local", action="store_true", help="使用内置离线演示页，不访问外网")
    parser.add_argument("--headless", action="store_true", help="无头模式运行，不弹出浏览器窗口")
    parser.add_argument(
        "--timeout", type=int, default=10, metavar="SEC", help="单个元素最长等待秒数（默认 10）"
    )
    parser.add_argument(
        "--page-load-timeout",
        type=int,
        default=30,
        metavar="SEC",
        help="页面加载最长等待秒数（默认 30）",
    )
    parser.add_argument("--keep-open", action="store_true", help="流程结束后保留浏览器窗口")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出 DEBUG 级日志")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    log_file = setup_logging(args.verbose)
    logger.debug("日志文件: %s", log_file)

    config = build_run_config(args)
    exit_code = run(config)
    logger.info("退出码: %d（0=成功 1=未知异常 2=超时 3=元素缺失 4=浏览器异常 130=手动中断）", exit_code)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
