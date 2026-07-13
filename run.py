#!/usr/bin/env python3
"""
自动读取默认 Excel 路径 → 生成 data.js → 打开看板
"""
import json
import tempfile
import webbrowser
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent
DATA_JS = BASE_DIR / "data.js"
DASHBOARD = BASE_DIR / "dashboard.html"

# Excel 源文件路径（Windows）
SRC = Path(
    r"C:\Users\wwang\OneDrive - Emerson\00 General - TM-CSC Calibration team\01 BorrowList&Standards list 2026.xlsx"
)


def safe_value(val):
    """把 NaN / NaT 转为 None，便于 JSON 序列化"""
    if pd.isna(val):
        return None
    if hasattr(val, "strftime"):
        # 日期类型 → ISO 字符串（不带时间，否则前端解析会有偏差）
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (bytes,)):
        return str(val)
    if isinstance(val, float) and val == int(val):
        return int(val)
    return val


def refresh_excel_file(path: Path) -> Path:
    """用本机 Excel 打开并刷新工作簿，返回可安全读取的文件快照路径。"""
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        print(f"⚠️ 无法使用 Excel 自动刷新（未安装 pywin32 或环境不支持）：{exc}")
        return path

    excel = None
    workbook = None
    temp_path = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        try:
            excel.AskToUpdateLinks = False
        except Exception:
            pass

        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        try:
            workbook.RefreshAll()
        except Exception:
            pass
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as tmp:
            temp_path = Path(tmp.name)
        try:
            workbook.SaveCopyAs(str(temp_path))
            return temp_path
        except Exception as exc:
            print(f"⚠️ Excel 已刷新，但生成快照失败，改为直接读取原文件：{exc}")
            if temp_path is not None and temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return path
    except Exception as exc:
        print(f"⚠️ Excel 刷新失败，继续读取本地文件：{exc}")
        return path
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main():
    if not SRC.exists():
        print(f"❌ 数据源不存在: {SRC}")
        print("   请确认 Windows 端 OneDrive 已同步。")
        return

    snapshot = refresh_excel_file(SRC)
    print(f"📂 读取: {SRC.name}")
    xl = pd.ExcelFile(snapshot)

    # 1. Standards List
    standards = []
    if "Standards List" in xl.sheet_names:
        df = xl.parse("Standards List")
        for _, row in df.iterrows():
            standards.append({
                "id": safe_value(row.get("ID")),
                "sn": safe_value(row.get("SN")),
                "model": safe_value(row.get("Model")),
                "manufacture": safe_value(row.get("Manufacture")),
                "description": safe_value(row.get("Description")),
                "calDate": safe_value(row.get("Cal Date")),
                "due": safe_value(row.get("Due")),
                "type": safe_value(row.get("Type")),
                "calVendor": safe_value(row.get("Cal Vendor")),
                "qty": safe_value(row.get("Qty")),
                "sendDate": safe_value(row.get("Send Date")),
                "matStatus": safe_value(row.get("MAT Status")),
                "active": safe_value(row.get("Active or Deactive")),
                "oversea": safe_value(row.get("Oversea Vendor")),
                "station": safe_value(row.get("Station_at")),
                "comments": safe_value(row.get("Comments")),
                "ci": safe_value(row.get("CI")),
            })
    print(f"   ✅ Standards List: {len(standards)} 条")

    # 2. Overseas Calibration（保留原始列名 → col0, col1, ...）
    overseas = []
    if "Overseas Calibration" in xl.sheet_names:
        df = xl.parse("Overseas Calibration")
        # 统一列名 col0, col1, ...
        df.columns = [f"col{i}" for i in range(len(df.columns))]
        for _, row in df.iterrows():
            overseas.append({k: safe_value(v) for k, v in row.items()})
    print(f"   ✅ Overseas Calibration: {len(overseas)} 条")

    # 3. 写入 data.js
    data = {
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "source": SRC.name,
        "standards": standards,
        "overseas": overseas,
    }
    js_content = "window.__AUTO_DATA__ = " + json.dumps(data, ensure_ascii=False, default=str) + ";"
    DATA_JS.write_text(js_content, encoding="utf-8")
    print(f"💾 已生成: {DATA_JS.name} ({DATA_JS.stat().st_size:,} 字节)")

    # 4. 打开看板
    webbrowser.open(DASHBOARD.as_uri())
    print(f"🌐 已打开: {DASHBOARD.name}")

    if snapshot != SRC:
        try:
            snapshot.unlink()
        except Exception:
            pass


if __name__ == "__main__":
    main()
