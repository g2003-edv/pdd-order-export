#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
拼多多商家后台 · 订单自动导出机器人
====================================
流程（与 SOP 一一对应）：
  1. 电商融「商户管理」搜店铺 → 打开「商户信息」→ 抓 店铺主账号手机号 + 店铺密码
  2. 拼多多商家后台「账号登录」→ 填账号密码 → 登录
  3. 触发短信验证时 → 自动去电商融「手机号管理/短信记录」按手机号筛选，
     抓最新一条【拼多多】验证码（校验手机号一致 + 10 分钟时效）→ 回填确认
  4. 发货管理/订单查询 → 下单时间=自定义 → 设 00:00:00 ~ 23:59:59 → 查询
  5. 批量导出 → 自定义报表 → 生成报表 → 已生成报表页轮询 → 下载 Excel
  6. 多店铺自动：退出当前账号 → 下一家

用法：
  python pdd_export.py                        # 按 config.yaml 的 shops 清单跑
  python pdd_export.py 示例食品专卖店          # 只跑指定店铺（可多个）
  python pdd_export.py --start 2026-09-22 --end 2026-09-27
  python pdd_export.py --login                # 强制重新人工登录电商融
  python pdd_export.py --manual-dates         # 日期选择器改自动为人工（拼多多页面改动时兜底）

说明：
  - 首次运行会弹出浏览器，请人工登录电商融系统一次，登录态保存在 .browser_profile/
  - 滑块验证等平台人工校验：有头模式（默认）下脚本会等待人工完成；无头模式直接报错退出
  - 店铺账号/密码从电商融页面实时读取；命令行模式不写明文，服务器前端可选择加密保存
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import subprocess
import sys
import time
import traceback
from pathlib import Path

import yaml
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PWTimeout,
)

BASE_DIR = Path(__file__).resolve().parent
DEBUG_DIR = Path(os.environ.get("PDD_DEBUG_DIR") or (BASE_DIR / "debug")).resolve()

DS_BASE = "https://inner.dianshangrong.com"
DS_MERCHANT_URL = DS_BASE + "/incoming-merchant/merchant"
DS_SMS_URL = DS_BASE + "/incoming-phoneNumber/phonecardsmsrecord"
PDD_BASE = "https://mms.pinduoduo.com"
PDD_LOGIN_URL = PDD_BASE + "/login"
PDD_ORDERS_URL = PDD_BASE + "/orders/list"
PDD_EXPORT_URL = PDD_BASE + "/orders/exportExcel?exportType=0"

MAX_BACKOFF_DAYS = 92  # 拼多多限制：可查时间窗 ≤ 3 个月（留 2 天余量）


def log(msg: str) -> None:
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def shot(page, name: str) -> str:
    """失败排查用截图，存 debug/ 目录。"""
    try:
        DEBUG_DIR.mkdir(exist_ok=True)
        path = DEBUG_DIR / f"{dt.datetime.now():%Y%m%d_%H%M%S}_{name}.png"
        page.screenshot(path=str(path), full_page=True)
        log(f"  📸 截图已保存: {path}")
        return str(path)
    except Exception:
        return ""


def human_pause(msg: str) -> None:
    input(f"\n⏸  {msg}\n    完成后回到本窗口按回车继续…\n")


# ---------------------------------------------------------------------------
# 电商融系统
# ---------------------------------------------------------------------------

def ensure_ds_login(page, force: bool = False, headless: bool = False,
                    creds: dict | None = None) -> None:
    """确保电商融已登录（登录态持久化在 .browser_profile）。

    未登录时：优先用表单填写的账密自动登录（creds）；
    自动登录遇到意外拦截或没给账密 → 等人在浏览器里手动完成。
    """
    log("打开电商融系统…")
    page.goto(DS_BASE, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)

    def logged_in() -> bool:
        try:
            return page.get_by_text("商户管理", exact=False).count() > 0
        except Exception:
            return False

    if not force and logged_in():
        log("电商融登录态有效 ✓")
        return

    if creds and creds.get("account") and creds.get("password"):
        try:
            _ds_auto_login(page, creds["account"], creds["password"])
            if logged_in():
                log("电商融自动登录成功 ✓（登录态已保存，下次免登录）")
                return
            log("  自动登录已提交但未检测到登录态，转人工兜底…")
        except Exception as e:
            log(f"  电商融自动登录未完成（{e}），转人工兜底…")

    if headless:
        raise RuntimeError(
            "电商融未登录。请用有头模式运行一次完成人工登录，或在表单里填写电商融账密"
        )

    log("电商融未登录：请在弹出的浏览器中完成登录（含任何系统验证码）…")
    deadline = time.time() + 600
    while time.time() < deadline:
        if logged_in():
            log("电商融登录成功 ✓（登录态已保存，下次免登录）")
            return
        page.wait_for_timeout(3000)
    raise RuntimeError("等待电商融登录超时（10 分钟）")


def _ds_solve_slider(page, max_attempts: int = 3) -> bool:
    """尝试自动通过电商融的 AJ-Captcha 拼图滑块。

    原理：截取背景面板图 + 抠出拼图块图片（带透明通道）→ cv2 模板匹配定位缺口
    → 拟人轨迹（缓动 + 抖动 + 过冲回正）拖动滑块。失败点刷新重试。
    任何一步失败都返回 False，由调用方回退到人工拖动。"""
    try:
        import base64
        import random

        import cv2
        import numpy as np
    except Exception:
        return False

    def visible(sel: str) -> bool:
        try:
            return page.locator(sel + ":visible").count() > 0
        except Exception:
            return False

    if not visible(".verify-move-block"):
        return False

    handle = page.locator(".verify-move-block").first
    panel = page.locator(".verify-img-panel").first
    piece = page.locator(".verify-sub-block").first

    for _attempt in range(max_attempts):
        try:
            p_box, h_box = panel.bounding_box(), handle.bounding_box()
            pc_box = piece.bounding_box()
            if not (p_box and h_box and pc_box):
                return False

            # 背景面板截图（缺口在背景图上）
            shot_path = DEBUG_DIR / f"slider_panel_{int(time.time())}.png"
            panel.screenshot(path=str(shot_path))
            bg = cv2.imread(str(shot_path))
            if bg is None:
                return False
            scale = p_box["width"] / bg.shape[1]  # 截图像素 -> CSS 像素

            # 拼图块：取 img src 的 base64 PNG（带 alpha），做带掩码的模板匹配
            piece_src = page.evaluate(
                """() => {
                    const img = document.querySelector('.verify-sub-block img')
                             || document.querySelector('.verify-sub-block');
                    if (!img) return '';
                    return img.tagName === 'IMG' ? (img.src || '') : '';
                }"""
            )
            gap_x_css = None
            if piece_src.startswith("data:image"):
                b64 = piece_src.split(",", 1)[1]
                arr = np.frombuffer(base64.b64decode(b64), np.uint8)
                tpl = cv2.imdecode(arr, cv2.IMREAD_UNCHANGED)  # BGRA
                if tpl is not None and tpl.shape[2] == 4:
                    mask = (tpl[:, :, 3] > 10).astype(np.uint8) * 255
                    tpl_bgr = tpl[:, :, :3]
                    # 面板截图比拼图块大：模板必须不大于背景图
                    h, w = tpl_bgr.shape[:2]
                    if h < bg.shape[0] and w < bg.shape[1]:
                        res = cv2.matchTemplate(bg, tpl_bgr, cv2.TM_CCORR_NORMED,
                                                mask=mask)
                        _, _, _, loc = cv2.minMaxLoc(res)
                        gap_x_css = loc[0] * scale
            if gap_x_css is None:
                # 兜底：把面板截图存下来供人工排查，靠人工拖
                return False

            piece_rel_x = pc_box["x"] - p_box["x"]   # 拼图块当前相对面板的位置
            dist = gap_x_css - piece_rel_x
            if dist <= 5:
                continue  # 识别结果可疑（缺口在起点），刷新重试

            # 拟人轨迹：从拖块中心按下，缓出加速 + 抖动，末端小幅过冲再回正
            sx = h_box["x"] + h_box["width"] / 2
            sy = h_box["y"] + h_box["height"] / 2
            page.mouse.move(sx, sy)
            page.mouse.down()
            steps = random.randint(28, 40)
            cur = sx
            for i in range(1, steps + 1):
                t = i / steps
                ease = 1 - (1 - t) ** 3
                cur = sx + dist * ease + random.uniform(-1.2, 1.2)
                page.mouse.move(cur, sy + random.uniform(-1.5, 1.5))
                page.wait_for_timeout(random.randint(8, 18))
            page.mouse.move(sx + dist + random.uniform(2, 4), sy)   # 过冲
            page.wait_for_timeout(random.randint(60, 140))
            page.mouse.move(sx + dist, sy, steps=3)                 # 回正
            page.wait_for_timeout(random.randint(80, 160))
            page.mouse.up()
            page.wait_for_timeout(1800)

            if not visible(".verifybox") or not visible(".verify-move-block"):
                return True  # 验证框消失 = 通过
            ok_msg = page.get_by_text(re.compile("验证成功|通过验证")).count() > 0
            if ok_msg:
                return True
            # 没过：刷新换一张再试
            try:
                page.locator(".verify-refresh").first.click(timeout=2000)
                page.wait_for_timeout(1200)
            except Exception:
                pass
        except Exception:
            return False
    return False


def _ds_auto_login(page, account: str, password: str) -> None:
    """电商融账密自动登录：填用户名密码 → 登录 → 滑块等人拖（最长 5 分钟）。"""
    log("电商融：用表单填写的账号自动登录…")
    if "/login" not in page.url:
        page.goto(DS_BASE + "/login", wait_until="domcontentloaded")
        page.wait_for_timeout(2500)
    user_box = page.get_by_placeholder("请输入用户名")
    user_box.wait_for(state="visible", timeout=10000)
    user_box.fill(account)
    pwd_box = page.get_by_placeholder("请输入密码")
    pwd_box.fill(password)
    if user_box.input_value().strip() != account:  # 防自动填充覆盖
        user_box.fill(account)
    log("电商融：提交登录…")
    page.get_by_role("button", name="登录", exact=True).first.click()
    page.wait_for_timeout(2500)
    slider_note = False
    deadline = time.time() + 300
    while time.time() < deadline:
        try:
            if page.get_by_text("商户管理", exact=False).count() > 0:
                return
        except Exception:
            pass
        slider_up = page.locator(".verify-move-block:visible").count() > 0
        if slider_up and not slider_note:
            slider_note = True   # 只提示一次，避免刷屏
            log("  电商融弹出滑块验证，自动拖动中…")
            try:
                if _ds_solve_slider(page):
                    log("  滑块自动验证通过 ✓")
                    slider_note = False
                    continue
                log("  滑块未通过：可在弹出的浏览器中手动拖一下，程序继续等待")
            except Exception:
                log("  滑块自动拖动异常：可手动拖一下，程序继续等待")
        elif not slider_up:
            slider_note = False
        page.wait_for_timeout(3000)
    raise RuntimeError("电商融自动登录等待超时（5 分钟）")


def _on_merchant_page(page) -> bool:
    """是否真的在商户管理列表页：URL 对且页面有商户页特有的「店铺类型」筛选项。

    首页也有同名「店铺名称」搜索框，会话过期后系统会把深链重定向到首页，
    只看 URL 或侧边栏文字都会误判（2026-09-29 早上踩坑）。
    """
    try:
        return (
            "incoming-merchant" in page.url
            and page.get_by_text("店铺类型", exact=True).count() > 0
        )
    except Exception:
        return False


def ds_goto_merchant_page(page, tries: int = 3) -> None:
    """确保落在商户管理列表页：goto 深链被跳回首页时，改点侧边栏菜单。

    注意：即使已在商户管理页也要强制重载一次——上一家店铺可能留下未关闭的
    商户信息弹窗，其遮罩层会拦截本店铺的搜索点击（复用会话路径不经过短信页，
    不会自然刷新页面，必须在此主动清场）。
    """
    for _ in range(tries):
        if _on_merchant_page(page):
            try:
                page.goto(DS_MERCHANT_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)
            except Exception:
                pass
            if _on_merchant_page(page):
                return
        log(f"  不在商户管理页（…{page.url[-35:]}），尝试进入…")
        try:
            page.goto(DS_MERCHANT_URL, wait_until="domcontentloaded")
        except Exception:
            pass
        page.wait_for_timeout(3500)
        if _on_merchant_page(page):
            return
        try:  # 菜单父项展开后子项同名，点最后一个（子项）
            items = page.get_by_text("商户管理", exact=True)
            if items.count() >= 2:
                items.last.click(timeout=4000)
            elif items.count() == 1:
                items.first.click(timeout=4000)
            page.wait_for_timeout(3500)
        except Exception:
            pass
    if not _on_merchant_page(page):
        shot(page, "ds_not_merchant_page")
        raise RuntimeError(
            "进不去商户管理页（系统一直跳回首页）：请在浏览器里人工点开左侧"
            "「商户管理→商户管理」，或重新运行 python pdd_export.py --login 人工登录一次"
        )


def ds_open_shop_modal(page, shop_name: str) -> str:
    """搜店铺并打开「商户信息」弹窗，返回结果行文本（用于状态校验）。"""
    log(f"电商融·商户管理：搜索店铺「{shop_name}」")
    ds_goto_merchant_page(page)
    page.wait_for_timeout(1500)

    box = page.get_by_placeholder("请输入店铺名称")
    box.click()
    box.fill(shop_name)
    page.get_by_role("button", name="搜索", exact=True).click()

    # 表格是服务端分页且较慢（数千条记录），轮询等结果行出现，不能固定 sleep
    row = None
    rows = None
    deadline = time.time() + 40
    while time.time() < deadline:
        rows = page.locator(".el-table__body tr").filter(has_text=shop_name)
        if rows.count() == 0:  # 兼容其他表格结构
            rows = page.locator("tr").filter(has_text=shop_name)
        if rows.count() > 0:
            row = rows.first
            break
        # 「无结果」只在加载动画消失后才可信（转圈时空态本来就挂着）
        loading = (
            page.locator(".el-loading-mask:visible").count() > 0
            or page.locator("[class*='loading']:visible").count() > 0
            or page.get_by_text("加载中").count() > 0
        )
        if not loading and (
            page.locator(".el-table__empty-text:visible").count() > 0
            or page.get_by_text(re.compile(r"共\s*0\s*条")).count() > 0
        ):
            break  # 加载完成且确实无结果
        page.wait_for_timeout(2000)

    if row is None:
        shot(page, "shop_not_found")
        raise RuntimeError(
            f"搜不到店铺「{shop_name}」：30s 内结果表未出现该店铺"
            "（若确认系统里有此店铺，请人工搜一次核对名称写法）"
        )
    if rows.count() > 1:
        log(f"  ⚠️ 店铺名命中 {rows.count()} 行（可能含固定列重复渲染），取第一行")
    row_text = re.sub(r"\s+", " ", row.text_content() or row.inner_text())
    if shop_name not in row_text:
        raise RuntimeError(f"命中行异常，请人工核对：{row_text[:80]}")
    # SOP 前提校验：店铺必须处于合作中
    if "店铺合作中" not in row_text:
        raise RuntimeError(f"店铺状态异常（非「店铺合作中」），暂不能登录导单。行信息：{row_text[:80]}")
    if "拼多多" not in row_text:
        log(f"  ⚠️ 该店铺平台标记不是拼多多，仍继续尝试：{row_text[:60]}")

    row.get_by_text("编辑", exact=True).first.click()
    page.get_by_text("商户信息", exact=True).first.wait_for(state="visible", timeout=15000)
    page.wait_for_timeout(500)
    log("  商户信息弹窗已打开 ✓")
    return row_text


def _form_value(page, label: str) -> str:
    """按表单项标签读取 input 值。

    页面上往往有两组同名字段（列表筛选表单 + 商户信息弹窗），
    所以优先在可见弹窗内找，且逐个候选只认「有值」的输入框。
    """

    def _scan(root) -> str:
        strategies = []
        # 1) Element UI 标准结构：label 精确匹配
        try:
            strategies.append(root.locator(".el-form-item").filter(
                has=root.locator(
                    "label.el-form-item__label",
                    has_text=re.compile(rf"^\s*{re.escape(label)}\s*$"),
                )
            ))
        except Exception:
            pass
        # 2) 表单项文本包含匹配
        try:
            strategies.append(root.locator(".el-form-item").filter(has_text=label))
        except Exception:
            pass
        for items in strategies:
            for i in range(min(items.count(), 8)):
                it = items.nth(i)
                try:
                    if not it.is_visible():
                        continue
                except Exception:
                    pass
                try:
                    v = it.locator("input").first.input_value(timeout=1500)
                except Exception:
                    continue
                if v and v.strip():
                    return v.strip()
        return ""

    # 优先：可见弹窗内取
    try:
        dlg = page.locator(".el-dialog:visible, [role='dialog']:visible")
        if dlg.count() > 0:
            v = _scan(dlg.first)
            if v:
                return v
    except Exception:
        pass
    # 其次：全页扫（空字段自动跳过）
    v = _scan(page)
    if v:
        return v
    # 3) xpath 兜底
    try:
        alt = page.locator(
            f"xpath=//*[normalize-space(text())='{label}']/following::input[1]"
        )
        return (alt.first.input_value(timeout=2000) or "").strip()
    except Exception:
        return ""


def _dump_form(page) -> None:
    """把弹窗内所有字段「标签 = 值」打到日志，用于诊断字段差异。"""
    try:
        rows = page.evaluate(
            """() => Array.from(document.querySelectorAll('.el-form-item'))
                  .map(it => {
                    const l = it.querySelector('.el-form-item__label, label');
                    const i = it.querySelector('input');
                    return (l ? l.textContent.trim() : '?') + ' = '
                           + (i ? (i.value || '(空)') : '(无输入框)');
                  })"""
        )
        rows = [r for r in rows if "(无输入框)" not in r]
        log("  弹窗字段快照：" + " | ".join(rows[:40]))
    except Exception as e:
        log(f"  字段快照读取失败：{e}")


def ds_read_credentials(page, wait_s: int = 15) -> tuple[str, str]:
    # 弹窗数据异步加载：轮询等主账号手机号出现
    phone = ""
    deadline = time.time() + wait_s
    while time.time() < deadline:
        phone = _form_value(page, "店铺主账号手机号")
        if re.fullmatch(r"1\d{10}", phone):
            break
        page.wait_for_timeout(1500)
    if not re.fullmatch(r"1\d{10}", phone):
        _dump_form(page)
        shot(page, "credentials_empty")
        raise RuntimeError(
            f"未能读到有效的店铺主账号手机号（读到 {phone!r}）："
            "可能该店铺资料未录入或字段名不同，见上方字段快照与 debug 截图"
        )
    password = _form_value(page, "店铺密码")
    if not password:
        _dump_form(page)
        raise RuntimeError("店铺密码为空：电商融资料不全，找维护人补录")
    log(f"  取到主账号手机号 {phone[:3]}****{phone[-4:]} 与店铺密码（长度{len(password)}）✓")
    return phone, password


def ds_close_modal(page) -> None:
    """关闭商户信息弹窗：点「关闭」后必须确认遮罩真的消失，否则按 ESC 兜底。

    遮罩残留会拦截下一家店铺的搜索点击（30s 超时连锁失败），所以这里不吞错。"""
    for _ in range(2):
        try:
            page.get_by_role("button", name="关闭", exact=True).first.click(timeout=3000)
            page.wait_for_timeout(800)
        except Exception:
            pass
        try:
            if page.locator(".el-overlay:visible").count() == 0:
                return
        except Exception:
            return
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(600)
        except Exception:
            pass
    try:
        if page.locator(".el-overlay:visible").count() > 0:
            log("  ⚠️ 商户信息弹窗未能确认关闭（下一家店铺会强制重载页面兜底）")
    except Exception:
        pass


def ds_fetch_sms_code(page, phone: str, after: dt.datetime, timeout_s: int) -> str:
    """短信记录页按手机号搜索，轮询抓最新【拼多多】验证码。"""
    log(f"电商融·短信记录：等待 {phone[:3]}****{phone[-4:]} 的拼多多验证码…")
    page.goto(DS_SMS_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    try:  # 全系统最新短信（不筛选），用于判断接码通道是否整体正常
        pre = page.locator(".el-table__body tr").first.locator("td").all_inner_texts()
        if len(pre) >= 3:
            log(f"  [短信页·全系统最新] {pre[0]} {pre[1]} {pre[2][:30]}")
    except Exception:
        pass

    box = page.get_by_placeholder("请输入手机号")
    box.click()
    box.fill(phone)

    deadline = time.time() + timeout_s
    poll = 0
    while time.time() < deadline:
        poll += 1
        try:
            page.get_by_role("button", name="搜索", exact=True).click()
        except Exception:
            pass
        page.wait_for_timeout(2500)
        rows = page.locator(".el-table__body tr")
        if poll % 3 == 1 and rows.count() > 0:  # 定期打印最新一条，便于判断短信是否进系统
            try:
                fc = rows.first.locator("td").all_inner_texts()
                if len(fc) >= 3:
                    log(f"  [短信页·最新] {fc[0]} {fc[1]} {fc[2][:36]}")
            except Exception:
                pass
        for i in range(min(rows.count(), 30)):
            cells = rows.nth(i).locator("td").all_inner_texts()
            if len(cells) < 3:
                continue
            recv_phone, recv_time_s, content = (c.strip() for c in cells[:3])
            if recv_phone != phone or "拼多多" not in content or "验证码" not in content:
                continue
            m = re.search(r"验证码[：:\s]*(\d{4,8})", content)
            if not m:
                continue
            try:
                recv_dt = dt.datetime.strptime(recv_time_s, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
            if recv_dt < after:
                continue  # 旧验证码，不能用
            log(f"  抓到验证码 {m.group(1)}（短信时间 {recv_time_s}）✓")
            return m.group(1)
        page.wait_for_timeout(5000)
    raise RuntimeError(
        f"{timeout_s}s 内未收到新验证码：请核对登录手机号是否一致，"
        "或在拼多多验证页点「重新发送」后重试"
    )


# ---------------------------------------------------------------------------
# 拼多多商家后台
# ---------------------------------------------------------------------------

def _wait_slider_manual(page, headless: bool, timeout_s: int = 900) -> None:
    """滑块/拼图为平台人工校验：有头模式下等人手动拖完（默认 15 分钟），无头模式直接失败。"""
    if headless:
        shot(page, "pdd_slider_headless")
        raise RuntimeError("拼多多弹出滑块/拼图验证，无头模式无法人工处理，请用有头模式")
    log("  出现滑块/拼图验证：请在浏览器中手动拖动完成（脚本等待中）…")
    deadline = time.time() + timeout_s
    while time.time() < deadline and _slider_visible(page):
        page.wait_for_timeout(2000)
    if _slider_visible(page):
        raise RuntimeError("等待滑块验证完成超时（5 分钟）")
    page.wait_for_timeout(1500)


def _pdd_logged_in(page) -> bool:
    try:
        return (
            "/login" not in page.url
            and page.get_by_text(re.compile(r"主账号")).count() > 0
        )
    except Exception:
        return False


def _slider_visible(page) -> bool:
    try:
        return (
            page.get_by_text(re.compile(r"拖动滑块|滑块验证|完成拼图")).count() > 0
            or page.locator("iframe[src*='captcha'], iframe[src*='verify']").count() > 0
        )
    except Exception:
        return False


def _sms_panel_visible(page) -> bool:
    try:
        return page.get_by_placeholder(re.compile("短信验证码")).count() > 0
    except Exception:
        return False


def pdd_ensure_logged_out(ctx, headless: bool) -> None:
    """开跑前/换店前确保拼多多无残留登录态。"""
    for pg in list(ctx.pages):
        if "pinduoduo.com" in pg.url and "/login" not in pg.url:
            log("检测到拼多多残留登录态，先退出…")
            try:
                pg.bring_to_front()
                pg.get_by_text(re.compile(r"主账号")).first.click(timeout=5000)
                pg.wait_for_timeout(800)
                pg.get_by_text("退出当前账号", exact=True).first.click(timeout=5000)
                pg.wait_for_timeout(1500)
                for name in ("确定", "确 定"):
                    btn = pg.get_by_role("button", name=name)
                    if btn.count() > 0:
                        btn.first.click(timeout=2000)
                        break
                pg.wait_for_timeout(2000)
            except Exception:
                log("  页面退出失败，改用清除拼多多 Cookie…")
                try:
                    ctx.clear_cookies(domain=".pinduoduo.com")
                except Exception:
                    if headless:
                        raise RuntimeError(
                            "无法退出拼多多残留账号（且当前无头模式），请用有头模式处理一次"
                        )
                    human_pause("请在浏览器里手动退出拼多多当前账号")
            return


def _wait_manual_code(timeout_s: int = 600) -> str | None:
    """人工兜底：等用户把验证码写进 debug/CODE.txt（适配后台/无交互运行）。"""
    code_file = DEBUG_DIR / "CODE.txt"
    try:
        DEBUG_DIR.mkdir(exist_ok=True)
        code_file.write_text("", encoding="utf-8")
    except Exception:
        pass
    log(f"  ⏳ 请在 {timeout_s}s 内把收到的短信验证码写入：{code_file}")
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            t = code_file.read_text(encoding="utf-8").strip()
            if re.fullmatch(r"\d{4,8}", t):
                log(f"  读取到人工验证码 {t} ✓")
                return t
        except Exception:
            pass
        time.sleep(2)
    return None


SESSION_SHOP_FILE = Path(
    os.environ.get("PDD_SESSION_FILE") or (DEBUG_DIR / "SESSION_SHOP.txt")
).resolve()


def _read_session_shop() -> str:
    try:
        return SESSION_SHOP_FILE.read_text(encoding="utf-8").strip()
    except Exception:
        return ""


def _write_session_shop(name: str) -> None:
    try:
        DEBUG_DIR.mkdir(exist_ok=True)
        SESSION_SHOP_FILE.write_text(name, encoding="utf-8")
    except Exception:
        pass


def _cdp_port() -> int:
    try:
        return int(os.environ.get("PDD_CDP_PORT", "9222"))
    except ValueError:
        return 9222


def _cdp_url() -> str:
    return f"http://127.0.0.1:{_cdp_port()}"


def _find_real_browser(cfg) -> str | None:
    """找本机真实 Chrome/Edge：config 可指定，否则探常见安装路径。"""
    custom = (cfg.get("browser") or {}).get("chrome_path")
    if custom and Path(custom).exists():
        return custom
    cands = (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        str(Path.home() / "AppData/Local/Google/Chrome/Application/chrome.exe"),
    )
    for c in cands:
        if Path(c).exists():
            return c
    return None


def pdd_cdp_browser(pw, cfg):
    """连接（必要时拉起）真实 Chrome/Edge 的 CDP 调试口，返回 Browser。

    真浏览器无自动化指纹，风控最低。用项目内独立 profile 目录：
    不动用户日常浏览器、可与其同时运行（且 Chrome 新版只允许在
    非默认 profile 上开调试口）。
    """
    url = _cdp_url()
    port = _cdp_port()
    try:
        b = pw.chromium.connect_over_cdp(url, timeout=4000)
        log(f"已连接到真实浏览器（CDP:{port}）✓")
        return b
    except Exception:
        pass
    exe = _find_real_browser(cfg)
    if not exe:
        raise RuntimeError(
            "未找到 Chrome/Edge：请安装 Chrome，或在 config.yaml 的 browser.chrome_path 指定路径"
        )
    prof_path = Path(os.environ.get("PDD_CDP_PROFILE_DIR") or
                     (BASE_DIR / "chrome_profile")).resolve()
    prof_path.mkdir(parents=True, exist_ok=True)
    prof = str(prof_path)
    log(f"调试口未开，正在拉起真实浏览器 {Path(exe).name}（端口 {port}，profile={prof}）…")
    subprocess.Popen(
        [exe, f"--remote-debugging-port={port}", f"--user-data-dir={prof}",
         "--no-first-run", "--no-default-browser-check", "--window-size=1400,920",
         "about:blank"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    deadline = time.time() + 40
    while time.time() < deadline:
        try:
            b = pw.chromium.connect_over_cdp(url, timeout=4000)
            log(f"真实浏览器已启动并连接（CDP:{port}）✓")
            return b
        except Exception:
            time.sleep(1.5)
    raise RuntimeError(
        f"真实浏览器 CDP:{port} 连接失败（40s），请检查端口是否被占用或浏览器是否正常启动"
    )


def pdd_ensure_logged_in(ctx, phone: str, password: str, ds_page, cfg, shop_name: str) -> object:
    """优先复用持久化会话：登录一次=消耗一次短信配额，能复用就不登录。

    会话与店铺的绑定记在 debug/SESSION_SHOP.txt（登录成功即写入，比页面文案更可靠）；
    绑定匹配或页面账号区含店铺名才复用，否则退出残留后走完整登录。
    """
    headless = cfg["browser"]["headless"]
    probe = ctx.new_page()
    log("拼多多：探测已有登录态（复用会话可避免触发短信）…")
    head = None
    try:
        probe.goto(PDD_ORDERS_URL, wait_until="domcontentloaded")
        probe.wait_for_timeout(4000)
        deadline = time.time() + 20
        while time.time() < deadline and head is None:
            if "/login" in probe.url:
                break
            try:
                el = probe.get_by_text(re.compile(r"主账号")).first
                head = re.sub(r"\s+", "", el.inner_text(timeout=3000)) or None
            except Exception:
                head = None
            if head is None:
                probe.wait_for_timeout(2000)
    except Exception as e:
        log(f"  探测登录态出错：{e}")

    bound = _read_session_shop()
    if head:
        if bound == shop_name or shop_name in head:
            log(f"拼多多登录态有效（{head[:40]}），复用会话，无需登录/短信 ✓")
            if bound != shop_name:
                _write_session_shop(shop_name)
            return probe
        log(f"  会话属于其他店铺（绑定={bound or '未知'}，页面={head[:30]}），需切换账号…")
    else:
        log("  无有效登录态（Cookie 失效或已被登出），需要重新登录…")
    try:
        probe.close()
    except Exception:
        pass
    pdd_ensure_logged_out(ctx, headless)
    page = pdd_login(ctx, phone, password, ds_page, cfg)
    _write_session_shop(shop_name)
    return page


def pdd_login(ctx, phone: str, password: str, ds_page, cfg) -> object:
    """拼多多账号密码登录 + 自动过短信验证，返回已登录页面。"""
    headless = cfg["browser"]["headless"]
    sms_timeout = cfg["pinduoduo"].get("sms_timeout_seconds", 180)
    page = ctx.new_page()
    log("拼多多：打开登录页…")
    page.goto(PDD_LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(4000)

    # 风控滑块可能在打开登录页时就出现：先等人拖完再继续
    if _slider_visible(page):
        _wait_slider_manual(page, headless)

    # 扫码登录模式：完全绕开短信验证码（适合短信被平台限流时）
    if cfg["pinduoduo"].get("qr_login"):
        log("扫码登录模式：请在浏览器中用「拼多多商家版App」扫码完成登录（最长等 10 分钟）…")
        qr_deadline = time.time() + 600
        while time.time() < qr_deadline:
            if _pdd_logged_in(page):
                log("扫码登录成功 ✓")
                return page
            if _slider_visible(page):
                _wait_slider_manual(page, headless)
            page.wait_for_timeout(3000)
        raise RuntimeError("扫码登录等待超时（5 分钟）")

    # 有时默认停在「扫码登录」页签：账密输入框不可见就先点「账号登录」
    acc = page.locator("input[placeholder*='账号名']:visible")
    if acc.count() == 0:
        try:
            page.get_by_text("账号登录", exact=True).first.click(timeout=6000)
            page.wait_for_timeout(1500)
        except Exception:
            pass
        acc = page.locator("input[placeholder*='账号名']:visible")
    if acc.count() == 0:
        shot(page, "pdd_login_page_unknown")
        try:
            body = page.locator("body").inner_text(timeout=3000)
            log("  [登录页文本] " + re.sub(r"[\s]+", " ", body[:200]))
        except Exception:
            pass
        raise RuntimeError(
            "拼多多登录页未见账密输入框：可能触发滑块/风控提示或页面改版，"
            "见 debug 截图与上方登录页文本"
        )
    acc.first.fill(phone)
    pwd = page.locator("input[type=password]:visible").last
    pwd.click(timeout=10000)
    pwd.fill(password)
    # 防浏览器自动填充覆盖：回读核对
    if acc.first.input_value().strip() != phone:
        acc.first.fill(phone)
    log(f"拼多多：以 {phone[:3]}****{phone[-4:]} 提交登录…")

    after = dt.datetime.now() - dt.timedelta(seconds=60)  # 允许 60s 时钟偏差
    page.get_by_role("button", name="登录", exact=True).click()

    deadline = time.time() + 300
    sms_rounds = 0
    last_state = ""
    while time.time() < deadline:
        page.wait_for_timeout(2500)
        if _pdd_logged_in(page):
            log("拼多多登录成功 ✓")
            return page
        state = f"url…{page.url[-45:]} 短信面板={_sms_panel_visible(page)} 滑块={_slider_visible(page)}"
        if state != last_state:
            log(f"  [拼多多页面状态] {state}")
            last_state = state
        if _slider_visible(page):
            _wait_slider_manual(page, headless)
            continue
        if _sms_panel_visible(page):
            sms_rounds += 1
            if sms_rounds > 3:
                raise RuntimeError("短信验证反复未通过（超过 3 轮），中止")
            # 面板不一定自动发码：有发送按钮就点（有倒计时说明已发过）
            try:
                if page.get_by_text(re.compile(r"还剩\s*\d+\s*秒")).count() == 0:
                    for t in ("获取验证码", "获取短信验证码", "重新发送", "重新获取"):
                        btn = page.get_by_text(t, exact=True)
                        if btn.count() > 0 and btn.first.is_visible():
                            btn.first.click()
                            log(f"  已点击「{t}」触发发送")
                            after = dt.datetime.now()
                            page.wait_for_timeout(1500)
                            break
            except Exception:
                pass
            try:  # 打印面板文字，便于诊断
                body = page.locator("body").inner_text(timeout=3000)
                k = body.find("账号安全")
                if k >= 0:
                    log(f"  [验证面板] " + re.sub(r"[\s]+", " ", body[k:k + 130]))
            except Exception:
                pass
            code = None
            send_clicks = 1   # 面板倒计时代表平台已发过第 1 次
            max_sends = 2     # 平台对验证码有频控：最多再主动重发 1 次，避免烧配额
            stage_deadline = time.time() + sms_timeout + 120
            while time.time() < stage_deadline and not code:
                if send_clicks < max_sends:
                    try:
                        if page.get_by_text(re.compile(r"还剩\s*\d+\s*秒")).count() == 0:
                            for t in ("重新获取", "重新发送", "获取验证码"):
                                btn = page.get_by_text(t, exact=True)
                                if btn.count() > 0 and btn.first.is_visible():
                                    btn.first.click()
                                    send_clicks += 1
                                    log(f"  已点击「{t}」重发短信（累计第{send_clicks}次）")
                                    after = dt.datetime.now()
                                    page.wait_for_timeout(2000)
                                    break
                    except Exception:
                        pass
                try:
                    code = ds_fetch_sms_code(ds_page, phone, after, 25)
                except RuntimeError:
                    code = None
            if not code:
                shot(page, "sms_auto_fail")
                log("  自动取码未成功，转人工兜底…")
                code = _wait_manual_code(600)
                if not code:
                    raise RuntimeError("自动+人工文件均未取得验证码")
            box = page.locator("input[placeholder*='短信验证码']:visible").first
            box.click()
            box.fill(code)
            page.get_by_role("button", name="确认", exact=True).click()
            page.wait_for_timeout(2500)
            after = dt.datetime.now()  # 下一轮只认更新的短信
            continue
        err = page.get_by_text(re.compile(r"账号或密码错误|密码错误|账号不存在"))
        if err.count() > 0:
            raise RuntimeError(
                "拼多多提示账号/密码错误：回电商融重新复制（注意空格），"
                "或找维护人核对系统内密码"
            )
    shot(page, "pdd_login_timeout")
    raise RuntimeError("拼多多登录 180s 未完成，详见 debug 截图")


def _close_popups(page) -> None:
    """尽力关掉消息/通知/活动类弹窗（ESC + 可见关闭图标 + 常见文案按钮）。

    弹窗盖住页面会挡住后续点击让流程看起来卡死，所以这里多轮清扫；
    只点“关掉不吃亏”的文案（我知道了/以后再说类），绝不碰「确定」，
    避免误确认危险操作。失败不影响主流程。"""
    dismiss_texts = ("我知道了", "知道了", "以后再说", "下次再说", "暂不升级",
                     "不再提示", "已知晓", "残忍拒绝", "残忍离开")
    for _ in range(3):
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(400)
        except Exception:
            pass
        clicked = 0
        try:
            for el in page.locator("[class*='close']:visible").all()[:4]:
                el.click(timeout=600)
                clicked += 1
                page.wait_for_timeout(200)
        except Exception:
            pass
        for txt in dismiss_texts:
            try:
                btn = page.get_by_text(txt, exact=True).first
                if btn.is_visible():
                    btn.click(timeout=600)
                    clicked += 1
                    page.wait_for_timeout(200)
            except Exception:
                pass
        if clicked == 0:
            break
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
    except Exception:
        pass


def _read_range_input(page) -> str:
    vals = page.evaluate(
        """() => Array.from(document.querySelectorAll('input'))
               .map(i => i.value || '')
               .filter(v => /\\d{4}-\\d{2}-\\d{2}/.test(v) && v.includes('~'))"""
    )
    return vals[0] if vals else ""


def _calendar_pick(page, start_dt: dt.datetime, end_dt: dt.datetime) -> None:
    """双日历面板选起止日期 + 时间归零/补 23:59:59。"""
    targets = [
        ((start_dt.year, start_dt.month), 0),
        ((end_dt.year, end_dt.month), 1),
    ]
    for (ty, tm), idx in targets:
        for _ in range(24):
            labels = page.locator(".el-date-picker__header-label").all_inner_texts()
            if len(labels) >= 4:
                year, mon = labels[idx * 2], labels[idx * 2 + 1]
                cur = (
                    int(re.search(r"(\d{4})", year).group(1)),
                    int(re.search(r"(\d{1,2})月", mon).group(1)),
                )
                if cur == (ty, tm):
                    break
                # 依次尝试该面板的左/右箭头，直到表头变化
                arrows = page.locator(".el-picker-panel__icon-btn")
                moved = False
                for ai in (idx * 2, idx * 2 + 1):
                    if ai >= arrows.count():
                        continue
                    try:
                        arrows.nth(ai).click(timeout=1500)
                        page.wait_for_timeout(600)
                        moved = True
                        break
                    except Exception:
                        continue
                if not moved:
                    raise RuntimeError("日期面板翻月失败")
            else:
                raise RuntimeError("未找到日期面板表头")
        else:
            raise RuntimeError("日期面板翻月超限")

        panel = page.locator(".el-date-table").nth(idx)
        cell = None
        for c in panel.locator("td.available").all():
            if c.inner_text().strip() == str((start_dt if idx == 0 else end_dt).day):
                cell = c
                break
        if cell is None:
            raise RuntimeError(f"日历中未找到 {ty}-{tm:02d} 的日期格")
        cell.click()
        page.wait_for_timeout(600)

    # 时间：先点「归零」(两个都变 00:00:00)，再把结束时间补到 23:59:59
    try:
        page.get_by_text("归零", exact=True).first.click(timeout=2500)
        page.wait_for_timeout(500)
    except Exception:
        pass
    tins = page.locator(".el-date-picker__time-header input")
    if tins.count() >= 2:
        end_in = tins.nth(1)
        try:
            end_in.click()
            end_in.fill("23:59:59")
            page.keyboard.press("Enter")
            page.wait_for_timeout(600)
        except Exception:
            # 时间滚轮面板：点 23 / 59 / 59
            for col, val in ((0, "23"), (1, "59"), (2, "59")):
                try:
                    page.locator(".el-time-spinner__list").nth(col).get_by_text(
                        val, exact=True
                    ).first.click(timeout=2000)
                except Exception:
                    pass
            try:
                page.locator(".el-time-panel__btn").get_by_text(
                    re.compile("确\\s*定")
                ).first.click(timeout=2000)
            except Exception:
                pass
    page.get_by_text(re.compile(r"^确\s*定$")).first.click()
    page.wait_for_timeout(800)


def _js_click_visible_text(page, text: str) -> bool:
    """在可见元素里按精确文本点击（拼多多自研组件无通用 role，用文本定位）。"""
    return page.evaluate(
        """(text) => {
            const els = [...document.querySelectorAll('div,span,li')]
                .filter(e => e.children.length===0 && (e.textContent||'').trim()===text
                    && e.offsetParent!==null);
            if (!els.length) return false;
            els[0].click();
            return true;
        }""",
        text,
    )


def _range_input_value(page) -> str:
    return page.evaluate(
        """() => {
            const i = document.querySelector(
                "input[data-testid='beast-core-rangePicker-htmlInput']");
            return i ? i.value : '';
        }"""
    )


def _pdd_pick_preset(page, preset: str, expect_date: str) -> None:
    """订单下单时间行的预设下拉（今日/昨日/近N日/自定义）选一项并校验生效。

    该行有两个同值控件：真正的下拉触发器是文本镜像 div（IPT_mirror），
    后面的 input 点了不展开；所以按「div 优先、input 兜底」逐个候选重试。
    行容器从 range 输入框向上找含「下单时间」的祖先，防止误点其他下拉。
    """
    def candidates_js():
        return page.evaluate(
            """() => {
                const inp = document.querySelector(
                    "input[data-testid='beast-core-rangePicker-htmlInput']");
                if (!inp) return [];
                let row = inp;
                for (let k = 0; k < 8 && row; k++) {
                    row = row.parentElement;
                    if (row && /下单时间/.test(row.textContent || '')) break;
                }
                if (!row) return [];
                const all = [...row.querySelectorAll('div,input')].filter(e =>
                    e.offsetParent !== null &&
                    /^(今日|昨日|近\\d+日|自定义)$/.test(
                        ((e.textContent || e.value) || '').trim()));
                all.sort((a, b) => (a.tagName === 'DIV' ? 0 : 1) - (b.tagName === 'DIV' ? 0 : 1));
                return all.map(e => ({
                    tag: e.tagName,
                    text: ((e.textContent || e.value) || '').trim(),
                }));
            }"""
        )

    cands = candidates_js()
    if not cands:
        raise RuntimeError("没找到下单时间预设下拉控件（页面里没有 今日/昨日/近N日/自定义）")
    log(f"  下单时间预设候选：{[(c['tag'], c['text']) for c in cands]}")
    last_err = "无候选可点"
    for idx in range(len(cands)):
        ok = page.evaluate(
            """(idx) => {
                const inp = document.querySelector(
                    "input[data-testid='beast-core-rangePicker-htmlInput']");
                if (!inp) return false;
                let row = inp;
                for (let k = 0; k < 8 && row; k++) {
                    row = row.parentElement;
                    if (row && /下单时间/.test(row.textContent || '')) break;
                }
                if (!row) return false;
                const all = [...row.querySelectorAll('div,input')].filter(e =>
                    e.offsetParent !== null &&
                    /^(今日|昨日|近\\d+日|自定义)$/.test(
                        ((e.textContent || e.value) || '').trim()));
                all.sort((a, b) => (a.tagName === 'DIV' ? 0 : 1) - (b.tagName === 'DIV' ? 0 : 1));
                if (!all[idx]) return false;
                all[idx].click();
                return true;
            }""",
            idx,
        )
        if not ok:
            continue
        page.wait_for_timeout(1200)
        if _js_click_visible_text(page, preset):
            page.wait_for_timeout(1500)
            page.keyboard.press("Escape")
            page.wait_for_timeout(600)
            val = _range_input_value(page)
            if expect_date in val or preset in val:
                log(f"  下单时间已设为预设「{preset}」：{val}")
                return
            last_err = f"点了候选但输入框未更新：{val!r}"
        else:
            last_err = "下拉未出现目标选项（点的是 input，或选项渲染慢）"
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
    raise RuntimeError(f"预设「{preset}」选择失败：{last_err}")


def _pdd_panel_open(page) -> bool:
    return page.evaluate(
        """() => {
            const root = document.querySelector(
                "[data-testid='beast-core-rangePicker-dropdown-contentRoot']");
            return !!root && root.getBoundingClientRect().width > 0;
        }"""
    )


def _pdd_open_panel(page) -> None:
    """点 range 输入框打开日历面板（有开/关 toggle 态，最多试 4 次）。"""
    for _ in range(4):
        page.locator(
            "input[data-testid='beast-core-rangePicker-htmlInput']").first.click()
        page.wait_for_timeout(1600)
        if _pdd_panel_open(page):
            return
    shot(page, "date_panel_closed")
    raise RuntimeError("日期面板未能打开")


def _pdd_labels(page):
    """两个月份面板的 [年, 月]，DOM 顺序=左右。"""
    return page.evaluate(
        r"""() => {
            const root = document.querySelector(
                "[data-testid='beast-core-rangePicker-dropdown-contentRoot']");
            if (!root) return [];
            return [...root.querySelectorAll('section')]
                .filter(s => s.querySelector('span.RPR_dateText_5-193-0'))
                .map(s => {
                    const y = s.querySelector(
                        "input[data-testid='beast-core-select-htmlInput']");
                    const m = s.querySelector('span.RPR_dateText_5-193-0');
                    return [
                        parseInt(((y && y.value) || '').replace(/[^0-9]/g, ''), 10) || 0,
                        parseInt(((m && m.textContent) || '').replace(/[^0-9]/g, ''), 10) || 0,
                    ];
                });
        }"""
    )


def _pdd_arrow(page, idx: int, direction: str) -> bool:
    ok = page.evaluate(
        r"""(args) => {
            const root = document.querySelector(
                "[data-testid='beast-core-rangePicker-dropdown-contentRoot']");
            if (!root) return false;
            const secs = [...root.querySelectorAll('section')]
                .filter(s => s.querySelector('span.RPR_dateText_5-193-0'));
            const sec = secs[args.idx];
            if (!sec) return false;
            const want = args.dir === 'prev'
                ? 'beast-core-icon-left' : 'beast-core-icon-right';
            const arr = [...sec.querySelectorAll('svg')]
                .find(s => (s.getAttribute('data-testid') || '') === want);
            if (!arr) return false;
            for (const ty of ['mousedown', 'mouseup', 'click'])
                arr.dispatchEvent(new MouseEvent(ty,
                    {bubbles: true, cancelable: true, view: window}));
            return true;
        }""",
        {"idx": idx, "dir": direction},
    )
    page.wait_for_timeout(700)
    return ok


def _pdd_navigate(page, idx: int, target: tuple) -> bool:
    """把第 idx 个面板导航到目标年月。被左右面板约束挡住时返回 False。"""
    stuck = 0
    for _ in range(36):
        labels = _pdd_labels(page)
        if idx >= len(labels):
            return False
        cur = tuple(labels[idx])
        if cur == tuple(target):
            return True
        direction = "next" if tuple(target) > cur else "prev"
        if not _pdd_arrow(page, idx, direction):
            stuck += 1
            if stuck >= 2:
                return False
        else:
            after = _pdd_labels(page)
            if idx < len(after) and tuple(after[idx]) == cur:
                stuck += 1
                if stuck >= 2:
                    return False
            else:
                stuck = 0
    return False


def _pdd_click_day(page, tbl_idx: int, day: int) -> None:
    ok = page.evaluate(
        r"""(args) => {
            const root = document.querySelector(
                "[data-testid='beast-core-rangePicker-dropdown-contentRoot']");
            if (!root) return 'NO_ROOT';
            const tables = [...root.querySelectorAll(
                "table[data-testid='beast-core-rangePicker-table']")];
            const tb = tables[args.tbl];
            if (!tb) return 'NO_TABLE';
            const cell = [...tb.querySelectorAll('.RPR_cell_5-193-0')].find(c =>
                (c.textContent || '').trim() === String(args.day) &&
                !String(c.className).includes('RPR_outOfMonth') &&
                !String(c.className).includes('RPR_disabled'));
            if (!cell) return 'NO_CELL';
            cell.click();
            return 'OK';
        }""",
        {"tbl": tbl_idx, "day": day},
    )
    if ok != "OK":
        raise RuntimeError(f"点日期 {day} 失败（{ok}）")
    page.wait_for_timeout(800)


def _pdd_times(page):
    return page.evaluate(
        r"""() => {
            const root = document.querySelector(
                "[data-testid='beast-core-rangePicker-dropdown-contentRoot']");
            if (!root) return {};
            const g = t => {
                const i = root.querySelector(
                    "input[data-testid='beast-core-rangePicker-timePicker-htmlInput-"
                    + t + "']");
                return i ? i.value : '';
            };
            return {begin: g('begin'), end: g('end')};
        }"""
    )


def _pdd_click_text_in_panel(page, text: str) -> bool:
    return page.evaluate(
        r"""(text) => {
            const root = document.querySelector(
                "[data-testid='beast-core-rangePicker-dropdown-contentRoot']");
            if (!root) return false;
            const el = [...root.querySelectorAll('*')].find(e =>
                (e.textContent || '').trim() === text && e.children.length === 0 &&
                e.getBoundingClientRect().width > 0);
            if (!el) return false;
            el.click();
            return true;
        }""",
        text,
    )


def _pdd_set_end_2359(page) -> None:
    """归零后两端都是 00:00:00；把结束时间滚轮设为 23:59:59。"""
    t = _pdd_times(page)
    if t.get("end") == "23:59:59":
        return
    end_inp = page.locator(
        "input[data-testid='beast-core-rangePicker-timePicker-htmlInput-end']").first
    try:
        end_inp.click(timeout=4000)
        page.wait_for_timeout(1500)
    except Exception as e:
        raise RuntimeError(f"结束时间输入框点不开：{e}")
    picked = page.evaluate(
        r"""() => {
            const pick = (testid, val) => {
                const ul = document.querySelector(
                    "ul[data-testid='beast-core-timePicker-list-" + testid + "']");
                if (!ul || ul.getBoundingClientRect().height === 0) return 'NO_UL';
                const li = [...ul.children].find(c =>
                    (c.textContent || '').trim() === val);
                if (!li) return 'NO_LI';
                li.click();
                return val;
            };
            return [pick('hh', '23'), pick('mm', '59'), pick('ss', '59')];
        }"""
    )
    page.wait_for_timeout(900)
    t = _pdd_times(page)
    if t.get("end") != "23:59:59":
        raise RuntimeError(f"结束时间设置未生效（{picked}，当前 {t}）")


def _pdd_pick_custom(page, sd: dt.date, ed: dt.date) -> None:
    """自定义日期段：开面板 → 导航月份 → 点起止日 → 归零 → 结束时间 23:59:59 → 确认。

    左右面板有「左<右」约束：某一面板到不了目标月份时自动换另一面板；
    起止同月时两次点击都落在同一个月视图里。
    """
    _pdd_open_panel(page)
    sm = (sd.year, sd.month)
    em = (ed.year, ed.month)
    if _pdd_navigate(page, 0, sm):
        start_tbl = 0
    elif _pdd_navigate(page, 1, sm):
        start_tbl = 1
    else:
        shot(page, "date_range_failed")
        raise RuntimeError(f"无法把面板导航到起始月份 {sm}")
    _pdd_click_day(page, start_tbl, sd.day)

    if sm == em:
        _pdd_click_day(page, start_tbl, ed.day)
    elif _pdd_navigate(page, 1, em):
        _pdd_click_day(page, 1, ed.day)
    elif _pdd_navigate(page, 0, em):
        _pdd_click_day(page, 0, ed.day)
    else:
        shot(page, "date_range_failed")
        raise RuntimeError(f"无法把面板导航到结束月份 {em}")

    if not _pdd_click_text_in_panel(page, "归零"):
        raise RuntimeError("没找到「归零」按钮")
    page.wait_for_timeout(600)
    t = _pdd_times(page)
    if t.get("begin") != "00:00:00":
        raise RuntimeError(f"归零未生效（{t}）")
    _pdd_set_end_2359(page)
    if not _pdd_click_text_in_panel(page, "确认"):
        raise RuntimeError("没找到「确认」按钮")
    page.wait_for_timeout(1200)
    got = _range_input_value(page)
    if f"{sd:%Y-%m-%d}" not in got or f"{ed:%Y-%m-%d}" not in got:
        shot(page, "date_range_failed")
        raise RuntimeError(f"自定义日期未生效（输入框={got!r}）")
    log(f"  下单时间已自定义：{got}")


def pdd_set_date_range(page, start_dt, end_dt, manual: bool) -> None:
    log(f"设置下单时间：{start_dt:%Y-%m-%d 00:00:00} ~ {end_dt:%Y-%m-%d 23:59:59}")
    _close_popups(page)

    if manual:
        human_pause(
            "请在浏览器中手动设置日期区间（起 00:00:00 ~ 止 23:59:59），"
            "确认输入框显示正确后"
        )
        return

    today = dt.date.today()
    if start_dt.date() == end_dt.date() == today - dt.timedelta(days=1):
        try:  # 例行导昨天：平台预设最稳
            _pdd_pick_preset(page, "昨日", f"{today - dt.timedelta(days=1):%Y-%m-%d}")
            return
        except Exception as e:
            log(f"  预设「昨日」失败，改走自定义面板：{e}")
    elif start_dt.date() == end_dt.date() == today:
        try:
            _pdd_pick_preset(page, "今日", f"{today:%Y-%m-%d}")
            return
        except Exception as e:
            log(f"  预设「今日」失败，改走自定义面板：{e}")

    try:
        _pdd_pick_custom(page, start_dt.date(), end_dt.date())
    except Exception as e:
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        shot(page, "date_range_failed")
        raise RuntimeError(f"自定义日期设置失败：{e}")


def pdd_search_orders(page, start_dt, end_dt, manual_dates: bool) -> int:
    log("打开 发货管理→订单查询 …")
    page.goto(PDD_ORDERS_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(4000)
    _close_popups(page)

    # 切到「全部订单」标签：页面默认停在「待发货」，只导未发货单会漏数据
    try:
        page.get_by_text("全部订单(近3个月)", exact=True).first.click(timeout=4000)
        page.wait_for_timeout(1500)
        log("  已切换到「全部订单」标签")
    except Exception:
        log("  ⚠️ 未找到「全部订单」标签，按当前标签继续")
    _close_popups(page)  # 换标签后可能有新弹窗，盖住筛选区会卡住后续点击

    pdd_set_date_range(page, start_dt, end_dt, manual_dates)
    _close_popups(page)  # 查询前再清一次，避免弹窗挡住「查询」按钮

    # 二次确认：页面初始化慢时，首次点击的标签可能被水合过程恢复成默认的
    # 「待发货」（实测同账号同日期出现 2561→1 单）。设置日期后再点一次标签
    # 并重设日期，保证最终查询落在「全部订单」+ 正确日期上。
    try:
        page.get_by_text("全部订单(近3个月)", exact=True).first.click(timeout=4000)
        page.wait_for_timeout(1500)
        pdd_set_date_range(page, start_dt, end_dt, manual_dates)
        log("  已二次确认「全部订单」标签与日期")
    except Exception:
        pass

    page.get_by_role("button", name="查询", exact=True).click()
    page.wait_for_timeout(3000)
    counter = page.get_by_text(re.compile(r"共查询到\s*([\d,]+)\s*个订单"))
    for _ in range(20):
        t = counter.first.inner_text() if counter.count() > 0 else ""
        m = re.search(r"共查询到\s*([\d,]+)", t or "")
        if m:
            n = int(m.group(1).replace(",", ""))
            log(f"  查询完成：共 {n} 个订单 ✓")
            return n
        page.wait_for_timeout(2000)
    shot(page, "search_count_missing")
    raise RuntimeError("未读到「共查询到 N 个订单」结果")


def _csv_to_xlsx(csv_path: Path, xlsx_path: Path) -> None:
    """把平台导出的 CSV 转成真 Excel 表格（openpyxl）。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "订单"
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        for row in csv.reader(f):
            ws.append(row)
    for col_idx, col in enumerate(ws.columns, 1):
        try:
            width = max(len(str(c.value or "")) for c in list(col)[:80])
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = \
                min(width + 2, 42)
        except Exception:
            pass
    wb.save(xlsx_path)


def pdd_export_and_download(ctx, page, start_dt, end_dt, shop_name: str,
                            out_dir: Path, timeout_s: int) -> Path:
    log("批量导出 → 生成报表 …")
    # 平台限制：两次生成报表间隔 ≥5 分钟（2026-09-29 实测确认，被拒不报错只是不生成）。
    # 先开报表页读最新一次生成时间，不足 5 分钟就等够——多店铺串行时必触发。
    probe = None
    try:
        probe = ctx.new_page()
        probe.goto(PDD_EXPORT_URL, wait_until="domcontentloaded")
        probe.wait_for_timeout(3000)
        gens = probe.evaluate(
            """() => {
                const t = document.body.innerText || '';
                return [...t.matchAll(
                    /生成报表时间[:：]\\s*(\\d{4}-\\d{2}-\\d{2})\\s+(\\d{2}:\\d{2})/g)]
                    .map(m => m[1] + ' ' + m[2]);
            }"""
        )
        if gens:
            newest = max(dt.datetime.strptime(g, "%Y-%m-%d %H:%M") for g in gens)
            elapsed = (dt.datetime.now() - newest).total_seconds()
            # 卡片时间只精确到分钟（最多少算 59s）：平台按实际生成时刻严格算 5 分钟，
            # 所以余量取 365s（实测 282s 时被静默拒绝）
            if elapsed < 365:
                wait_s = int(365 - elapsed) + 5
                log(f"  ⏳ 平台限制两次生成间隔≥5分钟（最新报表 {newest:%H:%M}），等待 {wait_s}s…")
                time.sleep(wait_s)
        else:
            log("  报表页暂无历史报表")
    except Exception as e:
        log(f"  报表间隔预检失败（忽略，继续尝试生成）：{e}")
    finally:
        if probe:
            try:
                probe.close()
            except Exception:
                pass

    t_click = dt.datetime.now()  # 报表时间只精确到分钟，比对时另行放宽 90s
    page.get_by_role("button", name="批量导出", exact=True).click()
    page.wait_for_timeout(2500)

    # 报表类型保持「自定义报表」（点一下保险）
    try:
        page.get_by_text("自定义报表", exact=True).first.click(timeout=3000)
        page.wait_for_timeout(500)
    except Exception:
        pass

    exp_page = None
    try:
        with ctx.expect_page(timeout=20000) as newp:
            page.get_by_text("生成报表", exact=True).first.click()
        exp_page = newp.value
        log("  已在新标签页打开「已生成报表」")
    except Exception:
        log("  未开新标签页，直接跳转报表页…")
        exp_page = page
        exp_page.goto(PDD_EXPORT_URL, wait_until="domcontentloaded")
    exp_page.wait_for_timeout(3000)
    _close_popups(exp_page)  # 报表页弹窗会挡住「下载报表」按钮

    s, e = f"{start_dt:%Y-%m-%d}", f"{end_dt:%Y-%m-%d}"
    log(f"  轮询等待报表生成（时间段 {s} ~ {e}）…")
    deadline = time.time() + timeout_s
    out = out_dir / f"{shop_name}_{s}_{e}.csv"
    min_gen = (t_click - dt.timedelta(seconds=90)).strftime("%Y-%m-%d %H:%M")
    while time.time() < deadline:
        exp_page.reload(wait_until="domcontentloaded")
        exp_page.wait_for_timeout(3000)
        # 原子式：扫描报表卡片（自研组件，无 table），范围+生成时间都匹配时
        # 在同一个 JS 调用里点「下载报表」——避免打标记被 React 重渲染抹掉的竞态
        try:
            with exp_page.expect_download(timeout=90000) as dl:
                matched = exp_page.evaluate(
                    r"""(args) => {
                        const btns = [...document.querySelectorAll('button,a')]
                            .filter(b => (b.textContent||'').trim() === '下载报表'
                                && b.getBoundingClientRect().width > 0);
                        for (const btn of btns) {
                            let card = btn;
                            for (let k = 0; k < 10 && card; k++) {
                                card = card.parentElement;
                                if (card && /拼单成功时间/.test(card.textContent || '')) break;
                            }
                            const t = card ? (card.textContent || '') : '';
                            const genM = t.match(
                                /生成报表时间[:：]\s*(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})/);
                            const rangeM = t.match(
                                /拼单成功时间\s*(\d{4}-\d{2}-\d{2})\s+\d{2}:\d{2}\s*到\s*(\d{4}-\d{2}-\d{2})/);
                            if (!rangeM || rangeM[1] !== args.s || rangeM[2] !== args.e)
                                continue;
                            const gen = genM ? (genM[1] + ' ' + genM[2]) : '';
                            if (gen && new Date(gen.replace(/-/g, '/')) <
                                new Date(args.minGen.replace(/-/g, '/')))
                                continue;  // 旧报表，不是本次生成的
                            btn.click();
                            return {gen: gen};
                        }
                        return null;
                    }""",
                    {"s": s, "e": e, "minGen": min_gen},
                )
        except Exception:
            # 点了下载但 90s 内没等到下载事件（报表仍在后台生成/链接未就绪）：
            # 不能直接放弃——等一会重新轮询，直到总时限（timeout_s）为止
            log("  点击下载后 90s 未出现下载事件，稍后重试…")
            exp_page.wait_for_timeout(10000)
            continue
        if not matched:
            exp_page.wait_for_timeout(8000)
            continue
        log(f"  找到本次报表（生成于 {matched.get('gen') or '未知'}），下载中…")
        try:
            dl.value.save_as(str(out))
        except Exception:
            # 下载事件偶发丢失：重扫一次并再点一次
            log("  首次下载未完成，重试一次…")
            with exp_page.expect_download(timeout=90000) as dl2:
                exp_page.evaluate(
                    r"""() => {
                        const b = [...document.querySelectorAll('button,a')]
                            .filter(x => (x.textContent||'').trim() === '下载报表'
                                && x.getBoundingClientRect().width > 0)[0];
                        if (b) b.click();
                        return !!b;
                    }"""
                )
            dl2.value.save_as(str(out))
        # 平台报表实为 CSV（UTF-8 BOM）：按真实格式命名（用户不需要转 Excel）
        with open(out, "rb") as f:
            head = f.read(2)
        if head.startswith(b"PK") and out.suffix == ".csv":
            real = out.with_suffix(".xlsx")
            out.rename(real)
            out = real
        elif not head.startswith(b"PK") and out.suffix == ".xlsx":
            real = out.with_suffix(".csv")
            out.rename(real)
            out = real
        size = out.stat().st_size
        if size < 1024:
            raise RuntimeError(f"下载文件异常过小（{size}B）：{out}")
        log(f"  报表已下载：{out}（{size/1024:.0f} KB）✓")
        try:
            exp_page.close()  # 报表标签页用完即关，避免标签越积越多
        except Exception:
            pass
        return out
    shot(exp_page, "report_timeout")
    raise RuntimeError(f"{timeout_s}s 内报表未生成（注意平台限制：两次导出需间隔 5 分钟以上）")


def pdd_logout(page) -> None:
    try:
        log("退出当前拼多多账号（换店）…")
        page.get_by_text(re.compile(r"主账号")).first.click(timeout=5000)
        page.wait_for_timeout(800)
        page.get_by_text("退出当前账号", exact=True).first.click(timeout=5000)
        page.wait_for_timeout(1500)
        for name in ("确定", "确 定"):
            btn = page.get_by_role("button", name=name)
            if btn.count() > 0:
                btn.first.click(timeout=2000)
                break
        page.wait_for_timeout(2000)
        if "/login" not in page.url:
            page.goto(PDD_LOGIN_URL, wait_until="domcontentloaded")
    except Exception as e:
        log(f"  ⚠️ 页面退出异常（{e}），尝试清除拼多多 Cookie")
        try:
            page.context.clear_cookies(domain=".pinduoduo.com")
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def resolve_shops(cfg, cli_shops):
    raw = cli_shops or cfg.get("shops") or []
    shops = []
    for it in raw:
        if isinstance(it, str):
            shops.append({"name": it})
        elif isinstance(it, dict) and it.get("name"):
            shops.append(it)
        else:
            raise ValueError(f"config.shops 条目不合法：{it!r}")
    if not shops:
        raise SystemExit(
            "未指定店铺：请在 config.yaml 的 shops 里配置清单，"
            "或在命令行传入店铺名，例如：python pdd_export.py 示例食品专卖店"
        )
    return shops


def resolve_dates(cfg, args):
    s = args.start or cfg.get("export", {}).get("start_date") or ""
    e = args.end or cfg.get("export", {}).get("end_date") or ""
    if args.days is not None and args.days >= 0:
        if s or e:
            raise SystemExit("--days 与 --start/--end 不要同时给")
        today = dt.date.today()
        s = (today - dt.timedelta(days=args.days)).isoformat()
        e = today.isoformat()
        log(f"按 --days {args.days} 导出：{s} ~ {e}（含今天）")
    if not s and not e:
        yday = dt.date.today() - dt.timedelta(days=1)
        s = e = yday.isoformat()
        log(f"未指定日期，默认导出昨天：{s}")
    elif not s or not e:
        raise SystemExit("--start 与 --end 必须成对给出（或都在 config.yaml 配好）")
    sd = dt.datetime.strptime(s, "%Y-%m-%d")
    ed = dt.datetime.strptime(e, "%Y-%m-%d").replace(hour=23, minute=59, second=59)
    if ed < sd:
        raise SystemExit("结束日期早于开始日期")
    if (ed - sd).days > MAX_BACKOFF_DAYS:
        log(f"⚠️ 时间窗超过 3 个月，拼多多会拒绝导出（当前 {(ed - sd).days} 天）")
    return sd, ed


def _kill_leftover_chrome(profile_dir: Path) -> None:
    """任务结束后关掉本任务拉起的真实浏览器，避免任务栏窗口越积越多。

    登录会话保存在 profile 磁盘文件里，关掉浏览器不影响下次免登录
    （下次任务会重新拉起并直接复用会话）。"""
    try:
        pat = str(profile_dir).replace("'", "''")
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
             "Where-Object { $_.CommandLine -like '*" + pat + "*' } | "
             "ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
             "-ErrorAction SilentlyContinue }"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=40)
    except Exception:
        pass


def _clear_tabs_keep_browser(ctx) -> None:
    """清掉自动化打开的标签页，但保留最后一个（导航到 about:blank）。

    标签页全部关闭会令真实浏览器因“没有窗口”而整体退出，后续店铺全部
    TargetClosedError——成功收尾和失败清理两条路径都必须用本函数。"""
    try:
        pgs = list(ctx.pages)
        for pg in pgs[:-1]:
            pg.close()
        if pgs:
            try:
                pgs[-1].goto("about:blank", wait_until="commit", timeout=8000)
            except Exception:
                pass
    except Exception:
        pass


def process_shop(ctx, ds_page, shop: dict, sd, ed, cfg, manual_dates: bool) -> dict:
    name = shop["name"]
    log(f"\n===== 店铺：{name} =====")
    if "account" in shop and "password" in shop:
        phone, password = shop["account"], shop["password"]
        log("  使用 config 内置账号密码（跳过电商融抓取）")
    else:
        ds_open_shop_modal(ds_page, name)
        phone, password = ds_read_credentials(ds_page)
        ds_close_modal(ds_page)

    page = pdd_ensure_logged_in(ctx, phone, password, ds_page, cfg, name)
    try:
        header = page.get_by_text(re.compile(r"主账号")).first.inner_text()
        if name not in header.replace(" ", ""):
            log(f"  ⚠️ 后台右上角账号信息与店铺名不完全一致，请人工确认：{header[:40]}")
    except Exception:
        pass

    n = pdd_search_orders(page, sd, ed, manual_dates)
    out_dir = Path(cfg["export"].get("output_dir") or "./downloads").resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    file = pdd_export_and_download(
        ctx, page, sd, ed, name, out_dir,
        cfg["pinduoduo"].get("report_timeout_seconds", 300),
    )
    # 跑完不登出：保留会话给下次任务用；但把自动化打开的标签页清掉（会话在 Cookie 里）。
    # 必须保留最后一个标签页（导航到 about:blank 清空内容）——若把标签页全部关闭，
    # 真实浏览器会因“没有窗口”而整体退出，多店铺任务的后续店铺全部报 TargetClosedError。
    _clear_tabs_keep_browser(ctx)
    return {"shop": name, "orders": n, "file": str(file), "status": "成功", "error": ""}


def main():
    # 日志编码统一：重定向到文件/webui 时固定 UTF-8（交互控制台保持系统编码）
    if hasattr(sys.stdout, "reconfigure") and not sys.stdout.isatty():
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    ap = argparse.ArgumentParser(description="拼多多商家后台订单自动导出机器人")
    ap.add_argument("shops", nargs="*", help="店铺名称（不填则用 config.yaml 的 shops 清单）")
    ap.add_argument("--config", default=str(BASE_DIR / "config.yaml"))
    ap.add_argument("--start", help="开始日期 YYYY-MM-DD（含当天 00:00:00）")
    ap.add_argument("--end", help="结束日期 YYYY-MM-DD（含当天 23:59:59）")
    ap.add_argument("--days", type=int, help="便捷用法：从 N 天前到今天（含今天），"
                                             "例如 --days 3 = 前天 00:00 ~ 今天 23:59")
    ap.add_argument("--headless", action="store_true", help="无头模式（弹滑块/未登录会直接失败）")
    ap.add_argument("--login", action="store_true", help="强制重新人工登录电商融")
    ap.add_argument("--ds-account", help="电商融账号（未登录时自动登录；也可用环境变量 DS_ACCOUNT）")
    ap.add_argument("--ds-password", help="电商融密码（也可用环境变量 DS_PASSWORD）")
    ap.add_argument("--out-dir", help="报表输出目录（默认用 config.yaml 的 export.output_dir）")
    ap.add_argument("--job-dir", help="服务器任务目录：追加写入 results.json")
    ap.add_argument("--profile-dir", help="电商融浏览器持久化目录（每位用户建议独立）")
    ap.add_argument("--cdp-profile-dir", help="拼多多真实浏览器 profile 目录（每位用户建议独立）")
    ap.add_argument("--cdp-port", type=int, help="拼多多真实浏览器 CDP 端口（每位用户建议独立）")
    ap.add_argument("--qr-login", action="store_true",
                    help="拼多多改用人工扫码登录（绕开短信验证码，适合短信被平台限流时）")
    ap.add_argument("--cdp", action="store_true",
                    help="拼多多走真实 Chrome/Edge（CDP 接管，风控最低；首次使用需人工登录一次）")
    ap.add_argument("--manual-dates", action="store_true", help="日期选择改为人工在浏览器里点")
    args = ap.parse_args()

    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8"))
    if args.headless:
        cfg["browser"]["headless"] = True
    if args.qr_login:
        cfg.setdefault("pinduoduo", {})["qr_login"] = True
    if args.out_dir:
        cfg.setdefault("export", {})["output_dir"] = args.out_dir
    if args.cdp_profile_dir:
        os.environ["PDD_CDP_PROFILE_DIR"] = str(Path(args.cdp_profile_dir).resolve())
    if args.cdp_port:
        os.environ["PDD_CDP_PORT"] = str(args.cdp_port)
    creds = {
        "account": args.ds_account or os.environ.get("DS_ACCOUNT", ""),
        "password": args.ds_password or os.environ.get("DS_PASSWORD", ""),
    }

    shops = resolve_shops(cfg, args.shops)
    sd, ed = resolve_dates(cfg, args)

    profile = Path(
        args.profile_dir or
        os.environ.get("PDD_DS_PROFILE_DIR") or
        (BASE_DIR / cfg["browser"].get("profile_dir", ".browser_profile"))
    ).resolve()
    profile.mkdir(parents=True, exist_ok=True)

    results = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(profile),
            headless=cfg["browser"]["headless"],
            viewport={"width": 1920, "height": 1080},
            locale="zh-CN",
            accept_downloads=True,
        )
        ds_page = ctx.pages[0] if ctx.pages else ctx.new_page()
        ensure_ds_login(ds_page, force=args.login, headless=cfg["browser"]["headless"],
                        creds=creds)

        # 拼多多部分默认用 Playwright 自带内核；--cdp 时走真实浏览器（风控最低）
        pdd_ctx = ctx
        cdp_browser = None
        if args.cdp:
            cdp_browser = pdd_cdp_browser(p, cfg)
            pdd_ctx = (cdp_browser.contexts[0] if cdp_browser.contexts
                       else cdp_browser.new_context())
            log("拼多多部分走真实浏览器（CDP 模式）")

        for shop in shops:
            try:
                results.append(
                    process_shop(pdd_ctx, ds_page, shop, sd, ed, cfg, args.manual_dates)
                )
            except Exception as e:
                try:  # 失败现场截图（最多前 4 页，防止标签堆积拖慢）
                    for pg in list(pdd_ctx.pages)[:4]:
                        shot(pg, "fail_dump")
                    for pg in list(ctx.pages)[:2]:
                        shot(pg, "fail_dump_ds")
                except Exception:
                    pass
                try:  # 失败也要保活浏览器：清标签页但保留最后一个（全关=浏览器退出）
                    _clear_tabs_keep_browser(pdd_ctx)
                except Exception:
                    pass
                log(f"❌ 店铺「{shop['name']}」失败：{e}")
                traceback.print_exc()
                results.append(
                    {"shop": shop["name"], "orders": "", "file": "",
                     "status": "失败", "error": str(e)[:200]}
                )
        ctx.close()
        if cdp_browser is not None:
            cdp_browser.close()  # 断开 CDP 连接
        # 主动关掉本任务拉起的真实浏览器：不留窗口在任务栏里。
        # 会话在 profile 磁盘文件中，下次任务重新拉起即可免登录复用。
        try:
            _kill_leftover_chrome(Path(
                os.environ.get("PDD_CDP_PROFILE_DIR")
                or (BASE_DIR / "chrome_profile")).resolve())
        except Exception:
            pass

    print("\n========== 运行汇总 ==========")
    for r in results:
        mark = "✅" if r["status"] == "成功" else "❌"
        print(f"{mark} {r['shop']}：{r['status']}"
              + (f"，{r['orders']} 单，{r['file']}" if r["file"] else f"（{r['error']}）"))

    if args.job_dir:
        try:
            job_dir = Path(args.job_dir).resolve()
            job_dir.mkdir(parents=True, exist_ok=True)
            (job_dir / "results.json").write_text(
                json.dumps(results, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:
            print(f"⚠️ 无法写入任务结果文件：{e}")


if __name__ == "__main__":
    main()
