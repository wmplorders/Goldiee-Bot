from __future__ import annotations

import asyncio
import threading

from .gmail_monitor import monitor_gmail
from .settings import PROJECT_ROOT
from .whatsapp_edge import WhatsAppEdge


async def run_bot(config: dict) -> None:
    from .tally_item_master import get_item_master

    try:
        item_master = await asyncio.to_thread(get_item_master, True)
        print(
            f"Item master warmed: {item_master['item_count']} items; "
            f"fingerprint {item_master['company_fingerprint']}."
        )
    except Exception as exc:
        print(f"Item master warm-up failed: {exc}")

    whatsapp = WhatsAppEdge(config)
    await whatsapp.start()
    loop = asyncio.get_running_loop()

    def send_gmail_result(message, attachments=()):
        future = asyncio.run_coroutine_threadsafe(
            whatsapp.send_to_chat(config["whatsapp"]["group_name"], message, attachments),
            loop,
        )
        future.result(timeout=180)

    def gmail_worker():
        try:
            monitor_gmail(config, PROJECT_ROOT, send_gmail_result)
        except Exception as exc:
            print(f"Gmail monitor is not running: {exc}")

    threading.Thread(target=gmail_worker, name="gmail-monitor", daemon=True).start()
    try:
        await whatsapp.listen_forever()
    finally:
        await whatsapp.close()
