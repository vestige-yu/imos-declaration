#!/usr/bin/env python3
import os
import socket
import sys
import threading
import traceback
import webbrowser
import zipfile
from xml.etree import ElementTree as ET

import app


def log_startup_error(message):
    try:
        app.DATA_DIR.mkdir(parents=True, exist_ok=True)
        log_path = app.DATA_DIR / "startup.log"
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(message.rstrip() + "\n\n")
    except Exception:
        pass


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def open_browser_and_wait(url):
    webbrowser.open(url)
    threading.Event().wait()


def self_test():
    preview = {
        "contractNo": "SELFTEST",
        "exportDateSerial": "",
        "consignee": "IMOS SELF TEST",
        "packageKind": "PALLET",
        "packageCount": 1,
        "grossWeight": 2.5,
        "netWeight": 2.0,
        "tradeTerm": "FOB",
        "originCountry": "中国",
        "destinationCountry": "意大利",
        "domesticSource": "苏州",
        "totals": {"quantity": 3, "amount": 12.34, "currency": "USD"},
        "commodityLines": [
            {
                "itemNo": "1",
                "hsCode": "8708999990",
                "goodsName": "汽车零部件",
                "quantity": 3,
                "amount": 12.34,
                "currency": "USD",
                "brand": "IMOS",
                "netWeight": 2.0,
            }
        ],
        "auditSamples": [
            {
                "itemNo": "1",
                "hsCode": "8708999990",
                "goodsName": "汽车零部件",
                "qadPartNo": "QAD-SELFTEST",
                "imosPartNo": "IMOS-SELFTEST",
                "invoiceQuantity": 3,
                "unitPrice": 4.113,
                "invoiceAmount": 12.34,
                "currency": "USD",
                "packingNetWeight": 2.0,
                "packingGrossWeight": 2.5,
                "poNo": "PO-SELFTEST",
                "invoiceSourceSheet": "Sheet 1",
                "invoiceSourceRow": 1,
            }
        ],
    }
    try:
        output_path, _ = app.generate_workbook(preview)
        with zipfile.ZipFile(output_path) as workbook:
            bad_file = workbook.testzip()
            if bad_file:
                raise RuntimeError(f"生成的 xlsx 压缩结构异常: {bad_file}")
            names = set(workbook.namelist())
            required = {
                "[Content_Types].xml",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/worksheets/sheet1.xml",
            }
            missing = required - names
            if missing:
                raise RuntimeError("生成的 xlsx 缺少必要文件: " + ", ".join(sorted(missing)))
            workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
            sheet_names = [
                item.attrib.get("name")
                for item in workbook_xml.findall(f"{{{app.NS_MAIN}}}sheets/{{{app.NS_MAIN}}}sheet")
            ]
            if "随机抽检" not in sheet_names:
                raise RuntimeError("生成的 xlsx 缺少随机抽检 sheet")
            for name in workbook.namelist():
                if name.endswith(".xml"):
                    ET.fromstring(workbook.read(name))
        print(f"SELF_TEST_OK {output_path}")
        return 0
    except Exception:
        log_startup_error("self-test failed:\n" + traceback.format_exc())
        traceback.print_exc()
        return 1


def main():
    if "--self-test" in sys.argv:
        return self_test()

    port = int(os.environ.get("PORT") or find_free_port())
    server, url = app.run_in_thread(host="127.0.0.1", port=port)
    try:
        try:
            import webview
        except ImportError:
            log_startup_error("pywebview unavailable, opened in the default browser.")
            open_browser_and_wait(url)
            return

        try:
            window = webview.create_window("IMOS 报关单生成", url, width=1280, height=860)
            webview.start()
            return window
        except Exception:
            log_startup_error("pywebview startup failed:\n" + traceback.format_exc())
            open_browser_and_wait(url)
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    raise SystemExit(main())
