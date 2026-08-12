"""打开群晖 QuickConnect，并按两步登录流程进入 DSM。"""

from __future__ import annotations

import getpass
import os
import time
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeoutError


NAS_URL = os.getenv("NAS_URL", "https://luyao-2023.cn5.quickconnect.cn/")
NAS_USERNAME = os.getenv("NAS_USERNAME", "陆遥015")
LOGIN_TIMEOUT_MS = int(os.getenv("NAS_LOGIN_TIMEOUT_MS", "60000"))
SEARCH_TIMEOUT_MS = int(os.getenv("NAS_SEARCH_TIMEOUT_MS", "120000"))
DOWNLOAD_TIMEOUT_MS = int(os.getenv("NAS_DOWNLOAD_TIMEOUT_MS", "600000"))
TARGET_FOLDER = os.getenv("NAS_TARGET_FOLDER", "FBM测款图")
DOWNLOAD_DIR = Path(os.getenv("NAS_DOWNLOAD_DIR", r"D:\NAS_download"))
SKCS_FILE = Path(os.getenv("NAS_SKCS_FILE", str(DOWNLOAD_DIR / "SKCS.txt")))


def get_password() -> str:
    """优先读取环境变量，否则在终端中隐藏输入密码。"""
    password = os.getenv("NAS_PASSWORD")
    if password:
        return password
    return 'Luyaousyyb25'


def load_skcs(path: Path) -> list[str]:
    """读取 SKC，一行一个；去除首尾空白、空行和重复项。"""
    if not path.is_file():
        raise FileNotFoundError(f"SKC 列表不存在：{path}")

    skcs: list[str] = []
    seen: set[str] = set()
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        skc = raw_line.strip()
        if not skc or skc in seen:
            continue
        if any(char in skc for char in '<>:"/\\|?*'):
            raise ValueError(f"SKC 含有 Windows 文件名非法字符：{skc}")
        seen.add(skc)
        skcs.append(skc)
    return skcs


def wait_for_visible_role(
    page: Page,
    role: str,
    name: str,
    *,
    timeout_ms: int,
) -> Locator:
    """在主页面和所有 iframe 中等待第一个可见的角色元素。"""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for frame in page.frames:
            locator = frame.get_by_role(role, name=name, exact=True)
            count = locator.count()
            for index in range(count):
                candidate = locator.nth(index)
                try:
                    if candidate.is_visible():
                        return candidate
                except PlaywrightError:
                    break
        page.wait_for_timeout(300)
    raise PlaywrightTimeoutError(
        f"等待元素超时：role={role!r}, name={name!r}"
    )


def login(page: Page, username: str, password: str) -> None:
    """完成 DSM 的用户名、密码两步登录。"""
    print(f"正在打开 NAS：{NAS_URL}", flush=True)
    page.goto(NAS_URL, wait_until="domcontentloaded", timeout=LOGIN_TIMEOUT_MS)

    # 若已有有效登录会话，DSM 桌面会直接出现。
    if page.get_by_role("button", name="显示桌面").count() == 1:
        print("当前会话已经登录。", flush=True)
        return

    username_input = page.get_by_role("textbox", name="用户帐号")
    username_input.wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)
    if username_input.count() != 1:
        raise RuntimeError("无法唯一定位“用户帐号”输入框。")

    next_button = page.get_by_role("button", name="登录")
    if next_button.count() != 1:
        raise RuntimeError("无法唯一定位用户名步骤的右箭头按钮。")

    username_input.fill(username)
    next_button.click()

    password_input = page.get_by_role("textbox", name="密码")
    password_input.wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)
    if password_input.count() != 1:
        raise RuntimeError("无法唯一定位“密码”输入框。")

    login_button = page.get_by_role("button", name="登录")
    if login_button.count() != 1:
        raise RuntimeError("无法唯一定位密码步骤的右箭头按钮。")

    password_input.fill(password)
    login_button.click()

    desktop_button = page.get_by_role("button", name="显示桌面")
    desktop_button.wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)
    print("登录成功，已进入 DSM 桌面。", flush=True)


def open_target_folder(page: Page, folder_name: str) -> None:
    """打开第一个 File Station，并点击指定文件夹。"""
    file_station_items = page.get_by_role(
        "menuitem", name="File Station", exact=True
    )
    file_station_items.first.wait_for(state="visible", timeout=LOGIN_TIMEOUT_MS)
    item_count = file_station_items.count()
    if item_count < 1:
        raise RuntimeError("DSM 桌面上没有找到 File Station。")

    # DSM 桌面上可能同时存在快捷方式和菜单项；按要求点击第一个。
    file_station_items.nth(0).click()
    print("已点击第一个 File Station，正在等待目标文件夹……", flush=True)

    deadline = time.monotonic() + LOGIN_TIMEOUT_MS / 1000
    while time.monotonic() < deadline:
        # File Station 通常运行在 DSM 创建的 iframe 中，因此逐个 frame 查找。
        for frame in page.frames:
            folder_items = frame.get_by_text(folder_name, exact=True)
            count = folder_items.count()
            for index in range(count):
                candidate = folder_items.nth(index)
                if candidate.is_visible():
                    candidate.dblclick()
                    print(f"已进入文件夹：{folder_name}", flush=True)
                    return
        page.wait_for_timeout(500)

    raise RuntimeError(f"等待文件夹“{folder_name}”出现超时。")


def download_skc(page: Page, skc: str, destination: Path) -> None:
    """在当前 File Station 文件夹中搜索一个 SKC，并保存为 ZIP。"""
    search_input = wait_for_visible_role(
        page, "textbox", "搜索", timeout_ms=LOGIN_TIMEOUT_MS
    )
    search_input.fill(skc)
    search_input.press("Enter")

    result = wait_for_visible_role(
        page, "option", skc, timeout_ms=SEARCH_TIMEOUT_MS
    )
    result.click()

    operation_button = wait_for_visible_role(
        page, "button", "操作", timeout_ms=LOGIN_TIMEOUT_MS
    )
    operation_button.click()
    download_item = wait_for_visible_role(
        page, "menuitem", "下载", timeout_ms=LOGIN_TIMEOUT_MS
    )

    temp_path = destination.with_name(destination.name + ".part")
    temp_path.unlink(missing_ok=True)
    print(f"[{skc}] 已找到，正在打包并下载……", flush=True)
    try:
        with page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as download_info:
            download_item.click()
        download = download_info.value
        download.save_as(temp_path)
        temp_path.replace(destination)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise
    print(f"[{skc}] 下载完成：{destination}", flush=True)


def download_missing_skcs(page: Page, skcs: list[str]) -> None:
    """跳过已有 ZIP，仅搜索和下载缺失的 SKC。"""
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    missing = [skc for skc in skcs if not (DOWNLOAD_DIR / f"{skc}.zip").is_file()]
    skipped = len(skcs) - len(missing)
    print(
        f"共读取 {len(skcs)} 个 SKC；已有 {skipped} 个，待下载 {len(missing)} 个。",
        flush=True,
    )

    failures: list[str] = []
    for index, skc in enumerate(missing, start=1):
        destination = DOWNLOAD_DIR / f"{skc}.zip"
        print(f"[{index}/{len(missing)}] 处理 {skc}", flush=True)
        try:
            download_skc(page, skc, destination)
        except (PlaywrightError, RuntimeError) as exc:
            failures.append(skc)
            print(f"[{skc}] 失败：{exc}", flush=True)

    if failures:
        raise RuntimeError("以下 SKC 下载失败：" + "、".join(failures))


def main() -> None:
    try:
        from cloakbrowser import launch
    except ImportError as exc:
        raise SystemExit(
            "缺少 cloakbrowser，请先在 sdsdiy_orders 使用的 Python 环境中运行此脚本。"
        ) from exc

    skcs = load_skcs(SKCS_FILE)
    if not skcs:
        print(f"SKC 列表为空：{SKCS_FILE}", flush=True)
        return

    missing = [skc for skc in skcs if not (DOWNLOAD_DIR / f"{skc}.zip").is_file()]
    if not missing:
        print("列表中的 ZIP 均已存在，无需打开 NAS。", flush=True)
        return

    password = get_password()
    browser = launch(
        headless=False,
        locale="zh-CN",
        timezone="Asia/Shanghai",
        args=["--start-maximized"],
    )
    context = browser.new_context(accept_downloads=True)

    try:
        page = context.pages[0] if context.pages else context.new_page()
        try:
            login(page, NAS_USERNAME, password)
            open_target_folder(page, TARGET_FOLDER)
            download_missing_skcs(page, skcs)
        except PlaywrightTimeoutError as exc:
            raise RuntimeError(
                "页面加载、登录确认或 File Station 打开超时。"
            ) from exc

        try:
            input("浏览完成后按回车关闭浏览器……")
        except EOFError:
            print("关闭浏览器窗口即可结束脚本。", flush=True)
            while context.pages:
                context.pages[0].wait_for_timeout(1000)
    except KeyboardInterrupt:
        print("\n收到退出指令，正在关闭浏览器……", flush=True)
    finally:
        try:
            context.close()
        except PlaywrightError:
            pass
        try:
            browser.close()
        except PlaywrightError:
            pass


if __name__ == "__main__":
    main()
