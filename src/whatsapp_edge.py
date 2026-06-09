from __future__ import annotations

import asyncio
import re
from collections import defaultdict, deque
from pathlib import Path
from typing import Iterable

from .command_router import recognized_command, run_command
from .state import load_state, save_state


class WhatsAppEdge:
    def __init__(self, config: dict):
        self.config = config
        self.context = None
        self.page = None
        self.playwright = None
        self.ui_lock = asyncio.Lock()
        self.initialized_chats: set[str] = set()
        self.recent_attachments: dict[str, deque[Path]] = defaultdict(lambda: deque(maxlen=10))

    async def start(self):
        try:
            from playwright.async_api import async_playwright
        except ImportError as exc:
            raise RuntimeError("Playwright is not installed. Run setup.bat first.") from exc

        self.playwright = await async_playwright().start()
        self.context = await self.playwright.chromium.launch_persistent_context(
            user_data_dir=self.config["whatsapp"]["edge_profile_dir"],
            executable_path=self.config["whatsapp"]["edge_executable"],
            headless=False,
            accept_downloads=True,
            downloads_path=self.config["paths"]["downloads"],
            args=["--disable-blink-features=AutomationControlled"],
        )
        self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
        await self.page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded")
        await self.wait_until_ready_or_qr()

    async def wait_until_ready_or_qr(self):
        assert self.page
        ready = "#pane-side, div[contenteditable='true'][role='textbox']"
        try:
            await self.page.wait_for_selector(ready, timeout=20000)
        except Exception:
            print("WhatsApp Web is not logged in. Scan the QR code in Edge; the bot will continue automatically.")
            await self.page.wait_for_selector(ready, timeout=0)

    async def open_chat(self, chat_name: str):
        assert self.page
        await self.wait_until_ready_or_qr()
        search = self.page.locator(
            "div[aria-label='Search input textbox'], "
            "div[title='Search input textbox'], "
            "#side div[contenteditable='true'][role='textbox']"
        ).first
        if not await search.count():
            search = self.page.locator("div[contenteditable='true'][role='textbox']").first
        await search.click()
        await search.fill(chat_name)
        await self.page.wait_for_timeout(800)
        result = self.page.locator(f"#pane-side span[title={_css_string(chat_name)}]").first
        if await result.count():
            await result.click()
        else:
            await self.page.keyboard.press("Enter")
        await self.page.wait_for_timeout(1200)

    async def current_chat_name(self) -> str:
        assert self.page
        title = self.page.locator("header span[title]").first
        return (await title.get_attribute("title") or "").strip() if await title.count() else ""

    async def send_to_chat(self, chat_name: str, text: str, attachments: Iterable[str | Path] = ()):
        async with self.ui_lock:
            await self._send_to_chat_unlocked(chat_name, text, attachments)

    async def _send_to_chat_unlocked(self, chat_name: str, text: str, attachments: Iterable[str | Path] = ()):
        await self.open_chat(chat_name)
        await self.send_message(text, attachments)

    async def send_message(self, text: str, attachments: Iterable[str | Path] = ()):
        assert self.page
        await self.wait_until_ready_or_qr()
        if text:
            box = self.page.locator("footer div[contenteditable='true'][role='textbox']").last
            await box.click()
            await box.fill(text)
            await self.page.keyboard.press("Enter")
        for attachment in attachments:
            await self.send_file(attachment)

    async def send_file(self, path: str | Path):
        assert self.page
        path = str(Path(path).resolve())
        attach = self.page.locator(
            "button[title='Attach'], button[aria-label='Attach'], span[data-icon='plus'], span[data-icon='clip']"
        ).first
        await attach.click()
        file_input = self.page.locator("input[type='file']").last
        await file_input.set_input_files(path)
        await self.page.wait_for_timeout(1000)
        await self.page.locator("span[data-icon='send']").last.click()
        await self.page.wait_for_timeout(1500)

    async def listen_forever(self):
        assert self.page
        poll_seconds = int(self.config.get("whatsapp_poll_seconds", 5))
        print("WhatsApp command listener is running.")
        while True:
            try:
                async with self.ui_lock:
                    await self._scan_unread_chats()
                    await self._scan_open_chat(process_existing=False)
            except Exception as exc:
                print(f"WhatsApp scan error: {exc}")
            await asyncio.sleep(poll_seconds)

    async def _scan_unread_chats(self):
        assert self.page
        unread = self.page.locator(
            "#pane-side span[data-icon='unread-count'], "
            "#pane-side [aria-label*='unread message' i]"
        )
        count = await unread.count()
        for index in range(min(count, 20)):
            marker = unread.nth(index)
            row = marker.locator("xpath=ancestor::div[@role='row'][1]")
            if not await row.count():
                row = marker.locator("xpath=ancestor::div[@role='listitem'][1]")
            if not await row.count():
                row = marker.locator("xpath=ancestor::div[contains(@class,'_ak8q')][1]")
            if await row.count():
                await row.click()
                await self.page.wait_for_timeout(700)
                await self._scan_open_chat(process_existing=True)

    async def _scan_open_chat(self, process_existing: bool = False):
        assert self.page
        chat_name = await self.current_chat_name()
        if not chat_name:
            return
        incoming = self.page.locator("div.message-in")
        count = await incoming.count()
        if chat_name not in self.initialized_chats and not process_existing:
            for index in range(max(0, count - 30), count):
                message = incoming.nth(index)
                self._mark_processed(await self._message_id(message, chat_name, index))
            self.initialized_chats.add(chat_name)
            return
        self.initialized_chats.add(chat_name)
        for index in range(max(0, count - 30), count):
            message = incoming.nth(index)
            message_id = await self._message_id(message, chat_name, index)
            if self._already_processed(message_id):
                continue

            files = await self._download_message_attachments(message, chat_name)
            for path in files:
                self.recent_attachments[chat_name].append(path)
            text = await self._message_text(message)
            if not recognized_command(text):
                self._mark_processed(message_id)
                continue

            command_files = files or list(self.recent_attachments[chat_name])
            try:
                result = await run_command(text, command_files, self.config)
                await self._send_to_chat_unlocked(chat_name, result.message, result.attachments)
                if text.casefold().strip() == "purchase summary" and result.attachments:
                    self.recent_attachments[chat_name].clear()
            except Exception as exc:
                await self._send_to_chat_unlocked(chat_name, f"Could not complete `{text.strip()}`: {exc}")
            self._mark_processed(message_id)

    async def _message_id(self, message, chat_name: str, index: int) -> str:
        message_id = await message.get_attribute("data-id")
        nested = message.locator("[data-id]").first
        if not message_id and await nested.count():
            message_id = await nested.get_attribute("data-id")
        if message_id:
            return message_id
        copyable = message.locator("[data-pre-plain-text]").first
        prefix = await copyable.get_attribute("data-pre-plain-text") if await copyable.count() else ""
        return f"{chat_name}:{prefix}:{index}"

    async def _message_text(self, message) -> str:
        selectable = message.locator("span.selectable-text").all_inner_texts()
        parts = await selectable
        if parts:
            return "\n".join(part.strip() for part in parts if part.strip())
        return (await message.inner_text()).strip()

    async def _download_message_attachments(self, message, chat_name: str) -> list[Path]:
        assert self.page
        buttons = message.locator(
            "button[aria-label*='Download' i], [title*='Download' i], span[data-icon='download']"
        )
        paths: list[Path] = []
        for index in range(await buttons.count()):
            button = buttons.nth(index)
            try:
                async with self.page.expect_download(timeout=10000) as download_info:
                    await button.click()
                download = await download_info.value
                safe_chat = re.sub(r'[<>:"/\\|?*]+', "-", chat_name)
                safe_name = re.sub(r'[<>:"/\\|?*]+', "-", download.suggested_filename)
                target = Path(self.config["paths"]["downloads"]) / "whatsapp" / safe_chat / safe_name
                target.parent.mkdir(parents=True, exist_ok=True)
                await download.save_as(target)
                paths.append(target)
            except Exception:
                continue
        return paths

    def _already_processed(self, message_id: str) -> bool:
        state = load_state(self.config["paths"]["state"])
        return message_id in state.get("processed_whatsapp_message_ids", [])

    def _mark_processed(self, message_id: str):
        state = load_state(self.config["paths"]["state"])
        ids = state.setdefault("processed_whatsapp_message_ids", [])
        if message_id not in ids:
            ids.append(message_id)
            state["processed_whatsapp_message_ids"] = ids[-2000:]
            save_state(self.config["paths"]["state"], state)

    async def close(self):
        if self.context:
            await self.context.close()
        if self.playwright:
            await self.playwright.stop()


def _css_string(value: str) -> str:
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def send_to_configured_group(config: dict, text: str, attachments: Iterable[str | Path] = ()):
    async def _run():
        whatsapp = WhatsAppEdge(config)
        await whatsapp.start()
        await whatsapp.send_to_chat(config["whatsapp"]["group_name"], text, attachments)

    asyncio.run(_run())


async def run_whatsapp_bot(config: dict):
    whatsapp = WhatsAppEdge(config)
    await whatsapp.start()
    try:
        await whatsapp.listen_forever()
    finally:
        await whatsapp.close()
