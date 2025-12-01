"""
Colab-ready script for fetching momo product titles from a CSV and exporting
an Excel summary.
"""
from __future__ import annotations

import io
import re
from datetime import date
from typing import Dict, List

import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    from google.colab import files  # type: ignore
except ImportError:  # pragma: no cover
    files = None  # type: ignore


def load_input_csv() -> pd.DataFrame:
    """Load the user-uploaded CSV into a DataFrame.

    Expects at least two columns: "序號" and "商品網址".
    """
    if files is None:
        raise RuntimeError("google.colab.files is required when running in Colab.")

    uploaded = files.upload()
    if not uploaded:
        raise ValueError("No file uploaded.")

    filename, data = next(iter(uploaded.items()))
    content = io.BytesIO(data)
    df = pd.read_csv(content)
    required_columns = {"序號", "商品網址"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")
    return df


def to_roc_date(target_date: date | None = None) -> str:
    """Convert a date to ROC format (YYY/MM/DD)."""
    target_date = target_date or date.today()
    roc_year = target_date.year - 1911
    return f"{roc_year}/{target_date.month}/{target_date.day}"


def _extract_title_from_soup(soup: BeautifulSoup) -> str:
    meta_title = soup.find("meta", property="og:title")
    if meta_title and meta_title.get("content"):
        return meta_title.get("content").strip()

    title_candidates = [
        ("h1", {"id": "osm_productName"}),
        ("h1", {"id": "goodsName"}),
        ("h1", {}),
        ("title", {}),
    ]
    for tag, attrs in title_candidates:
        found = soup.find(tag, attrs=attrs)
        if found and found.get_text(strip=True):
            return found.get_text(strip=True)

    return ""


def fetch_momo_product(url: str, timeout: int = 10, max_retries: int = 3) -> str:
    """Fetch the product title from a momo product URL.

    Returns an empty string on failure.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }

    for _ in range(max_retries):
        try:
            response = requests.get(url, headers=headers, timeout=timeout)
            if response.status_code != 200:
                continue
            soup = BeautifulSoup(response.text, "html.parser")
            title = _extract_title_from_soup(soup)
            if title:
                return title
        except requests.RequestException:
            continue
    return ""


def build_output_rows(df: pd.DataFrame, roc_date: str) -> List[Dict[str, str]]:
    columns = [
        "編號",
        "檢查案號",
        "查核日期",
        "網路名稱/店家名稱",
        "賣家帳號或拍賣代碼",
        "商品名稱",
        "再查核日期",
        "是否下架",
        "是否改正",
        "調查結果",
        "網址/地址",
        "商檢標識",
        "已宣導",
        "已下架",
    ]

    rows: List[Dict[str, str]] = []
    for _, row in df.iterrows():
        product_url = str(row.get("商品網址", "")).strip()
        product_title = fetch_momo_product(product_url) if product_url else ""
        rows.append(
            {
                "編號": row.get("序號", ""),
                "檢查案號": "",
                "查核日期": roc_date,
                "網路名稱/店家名稱": "momo購物網",
                "賣家帳號或拍賣代碼": "",
                "商品名稱": product_title,
                "再查核日期": "",
                "是否下架": "",
                "是否改正": "",
                "調查結果": "",
                "網址/地址": product_url,
                "商檢標識": "",
                "已宣導": "",
                "已下架": "",
            }
        )
    return rows


def export_to_excel(rows: List[Dict[str, str]], roc_date: str) -> str:
    columns = [
        "編號",
        "檢查案號",
        "查核日期",
        "網路名稱/店家名稱",
        "賣家帳號或拍賣代碼",
        "商品名稱",
        "再查核日期",
        "是否下架",
        "是否改正",
        "調查結果",
        "網址/地址",
        "商檢標識",
        "已宣導",
        "已下架",
    ]
    df_out = pd.DataFrame(rows, columns=columns)
    date_digits = re.sub(r"[^0-9]", "", roc_date)
    filename = f"momo_check_output_{date_digits}.xlsx"
    df_out.to_excel(filename, index=False)

    if files is not None:
        files.download(filename)
    return filename


def main() -> None:
    df_input = load_input_csv()
    roc_today = to_roc_date()
    rows = build_output_rows(df_input, roc_today)
    exported = export_to_excel(rows, roc_today)
    print(f"Exported: {exported}")


if __name__ == "__main__":
    main()
