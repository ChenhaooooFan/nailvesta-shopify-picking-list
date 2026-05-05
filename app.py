import io
import re
import unicodedata
from typing import Dict, List, Tuple

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st


# =========================
# 基础清洗/识别函数
# =========================
SIZE_SET = {"S", "M", "L"}
SPECIAL_KEYWORDS = [
    "choose", "free giveaway", "giveaway", "binder", "organizer", "toolkit", "toolkits",
    "no nails", "mixed", "混合", "赠品", "工具", "收纳"
]


def norm_text(x) -> str:
    """统一大小写、空格、特殊符号，方便匹配产品名。"""
    if pd.isna(x):
        return ""
    s = str(x)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("’", "'").replace("‘", "'").replace("`", "'")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def clean_cell(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def find_col(df: pd.DataFrame, candidates: List[str]) -> str:
    """按多个可能列名找列；支持大小写/空格差异。"""
    normalized_cols = {norm_text(c): c for c in df.columns}
    for c in candidates:
        key = norm_text(c)
        if key in normalized_cols:
            return normalized_cols[key]
    # 模糊包含匹配
    for c in candidates:
        key = norm_text(c)
        for nk, real in normalized_cols.items():
            if key and key in nk:
                return real
    return ""


def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    """读取产品图册：支持 CSV / XLSX。"""
    name = uploaded_file.name.lower()
    data = uploaded_file.read()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(data))

    # CSV 尝试多种常见编码
    for enc in ["utf-8-sig", "utf-8", "gb18030", "latin1"]:
        try:
            return pd.read_csv(io.BytesIO(data), encoding=enc)
        except Exception:
            continue
    raise ValueError("产品图册无法读取，请上传 CSV 或 Excel 文件。")


def build_catalog_maps(catalog_df: pd.DataFrame) -> Tuple[Dict[str, dict], Dict[str, dict]]:
    """从产品图册生成 SKU 和产品名映射。"""
    sku_col = find_col(catalog_df, ["SKU", "sku", "Variant SKU", "变体SKU"])
    name_col = find_col(catalog_df, ["款式英文名称", "Product Name", "产品英文名", "英文名称", "Title", "选项引用列"])
    loc_col = find_col(catalog_df, ["库位", "Location", "Bin", "货位"])

    missing = []
    if not sku_col:
        missing.append("SKU")
    if not name_col:
        missing.append("款式英文名称 / Product Name")
    if not loc_col:
        missing.append("库位 / Location")
    if missing:
        raise ValueError("产品图册缺少必要列：" + "、".join(missing))

    sku_map = {}
    name_map = {}

    for _, r in catalog_df.iterrows():
        sku = clean_cell(r.get(sku_col, ""))
        name = clean_cell(r.get(name_col, ""))
        loc = clean_cell(r.get(loc_col, ""))

        if sku:
            # 图册通常是基础 SKU；slip 里通常是 SKU-S/M/L
            base_sku = re.sub(r"-(S|M|L)$", "", sku.strip(), flags=re.I).upper()
            sku_map[base_sku] = {"product_name": name, "location": loc, "catalog_sku": base_sku}

        if name:
            name_map[norm_text(name)] = {"product_name": name, "location": loc, "catalog_sku": sku}

    return sku_map, name_map


def looks_like_sku(line: str) -> bool:
    """识别 slip 里的 SKU 行，例如 NPF015-L / NB001。"""
    s = line.strip()
    if not s:
        return False
    # 排除明显不是 SKU 的行
    if re.search(r"\s", s):
        return False
    # 常见 NailVesta SKU：NPJ021-L, NOF031-L, NB001
    return bool(re.match(r"^[A-Z]{1,5}[A-Z0-9]{2,8}(?:-[SML])?$", s, flags=re.I))


def parse_qty(line: str) -> int:
    """Shopify slip 常见格式是 1 of 1 / 2 of 2；取前面的数量。"""
    m = re.search(r"(\d+)\s+of\s+\d+", line.strip(), flags=re.I)
    if m:
        return int(m.group(1))
    m = re.search(r"\d+", line.strip())
    if m:
        return int(m.group(0))
    return 1


def extract_order_no(lines: List[str]) -> str:
    for line in lines:
        m = re.search(r"Order\s*#\s*([A-Z]*\d+)", line, flags=re.I)
        if m:
            return "#" + m.group(1).lstrip("#")
    return ""


def extract_items_block(lines: List[str]) -> List[str]:
    start = None
    for idx, line in enumerate(lines):
        txt = norm_text(line)
        # 有些 PDF 会把表头抽成一行 "ITEMS QUANTITY"，有些会抽成两行 "ITEMS" / "QUANTITY"
        if txt in {"items quantity", "items quantity:"} or ("ITEMS" in line.upper() and "QUANTITY" in line.upper()):
            start = idx + 1
            break
        if txt == "items" and idx + 1 < len(lines) and norm_text(lines[idx + 1]) == "quantity":
            start = idx + 2
            break
    if start is None:
        return []

    end = len(lines)
    for idx in range(start, len(lines)):
        txt = norm_text(lines[idx])
        if txt.startswith("thank you") or txt == "nailvesta" or "support@nailvesta.com" in txt or "nailvesta.com" in txt:
            end = idx
            break
    return [x.strip() for x in lines[start:end] if x.strip()]


def parse_slip_pdf(uploaded_pdf) -> pd.DataFrame:
    """解析独立站 packing slip PDF，输出订单明细。"""
    pdf_bytes = uploaded_pdf.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    rows = []

    for page_index, page in enumerate(doc, start=1):
        text = page.get_text("text")
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        order_no = extract_order_no(lines)
        block = extract_items_block(lines)

        i = 0
        while i < len(block):
            # 找到下一行 SKU
            sku_idx = None
            for j in range(i, len(block)):
                if looks_like_sku(block[j]):
                    sku_idx = j
                    break

            if sku_idx is None:
                break

            product_parts = block[i:sku_idx]
            full_sku = block[sku_idx].strip().upper()
            qty = 1
            next_i = sku_idx + 1
            if sku_idx + 1 < len(block):
                if re.search(r"\d+\s+of\s+\d+", block[sku_idx + 1], flags=re.I):
                    qty = parse_qty(block[sku_idx + 1])
                    next_i = sku_idx + 2

            size = ""
            if product_parts and product_parts[-1].strip().upper() in SIZE_SET:
                size = product_parts[-1].strip().upper()
                product_name = " ".join(product_parts[:-1]).strip()
            else:
                m = re.search(r"-(S|M|L)$", full_sku, flags=re.I)
                if m:
                    size = m.group(1).upper()
                product_name = " ".join(product_parts).strip()

            base_sku = re.sub(r"-(S|M|L)$", "", full_sku, flags=re.I).upper()

            if product_name or full_sku:
                rows.append({
                    "Order #": order_no,
                    "Page": page_index,
                    "Slip Product Name": product_name,
                    "Full SKU": full_sku,
                    "Base SKU": base_sku,
                    "Size": size,
                    "Qty": qty,
                })

            i = next_i

    if not rows:
        raise ValueError("没有从 slip PDF 里识别到商品。请确认上传的是独立站 Packing Slip PDF。")
    return pd.DataFrame(rows)


def is_special_item(product_name: str, sku: str, size: str) -> bool:
    txt = norm_text(product_name + " " + sku)
    if not size:
        return True
    return any(k in txt for k in SPECIAL_KEYWORDS)


def enrich_and_summarize(details_df: pd.DataFrame, sku_map: Dict[str, dict], name_map: Dict[str, dict]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    enriched = details_df.copy()
    product_names = []
    locations = []
    match_types = []

    for _, r in enriched.iterrows():
        base_sku = clean_cell(r.get("Base SKU", "")).upper()
        slip_name = clean_cell(r.get("Slip Product Name", ""))
        size = clean_cell(r.get("Size", "")).upper()

        match = None
        match_type = "未匹配"
        name_key = norm_text(slip_name)

        # 第一优先级：按 slip 里的 Base SKU 匹配产品图册。
        # 原因：仓库检货应以 SKU 为准；slip 上的款式名只作为 SKU 找不到时的兜底。
        if base_sku and base_sku in sku_map:
            match = sku_map[base_sku]
            catalog_name = clean_cell(match.get("product_name", ""))
            if slip_name and norm_text(catalog_name) and norm_text(catalog_name) != name_key:
                match_type = "SKU匹配(名称不一致)"
            else:
                match_type = "SKU匹配"
        elif name_key and name_key in name_map:
            match = name_map[name_key]
            match_type = "款式名兜底匹配"

        if match:
            product_name = match.get("product_name") or slip_name
            location = match.get("location") or ""
        else:
            product_name = slip_name or base_sku
            location = ""

        if not location:
            location = "无库位(特殊款)" if is_special_item(product_name, base_sku, size) else "未识别库位"

        product_names.append(product_name)
        locations.append(location)
        match_types.append(match_type)

    enriched["Product Name"] = product_names
    enriched["库位"] = locations
    enriched["匹配方式"] = match_types

    # 主检货单：和示例第四个文件一致
    summary_rows = []
    grouped = enriched.groupby(["库位", "Product Name"], dropna=False)
    for (loc, pname), g in grouped:
        s_qty = int(g.loc[g["Size"].str.upper() == "S", "Qty"].sum())
        m_qty = int(g.loc[g["Size"].str.upper() == "M", "Qty"].sum())
        l_qty = int(g.loc[g["Size"].str.upper() == "L", "Qty"].sum())
        total = int(g["Qty"].sum())
        summary_rows.append({"库位": loc, "Product Name": pname, "S": s_qty, "M": m_qty, "L": l_qty, "Total": total})

    summary = pd.DataFrame(summary_rows)
    if summary.empty:
        summary = pd.DataFrame(columns=["库位", "Product Name", "S", "M", "L", "Total"])

    summary = summary.sort_values(by="库位", key=lambda col: col.map(location_sort_key)).reset_index(drop=True)
    return summary, enriched


def location_sort_key(loc: str):
    s = clean_cell(loc)
    m = re.match(r"^([A-Za-z]+)-(\d+)-(\d+)$", s)
    if m:
        prefix, a, b = m.groups()
        return (0, prefix.upper(), int(a), int(b), s)
    if s == "未识别库位":
        return (8, s, 999, 999, s)
    if s.startswith("无库位"):
        return (9, s, 999, 999, s)
    return (7, s, 999, 999, s)


def make_excel(summary_df: pd.DataFrame, detail_df: pd.DataFrame) -> bytes:
    """生成带格式的 Excel。"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="检货单")
        detail_df.to_excel(writer, index=False, sheet_name="订单明细")

        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter"
        })
        cell_fmt = workbook.add_format({"border": 1, "valign": "vcenter"})
        center_fmt = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        warn_fmt = workbook.add_format({"bg_color": "#FFF2CC"})
        total_fmt = workbook.add_format({"border": 1, "bold": True, "align": "center", "valign": "vcenter"})

        ws = writer.sheets["检货单"]
        for col_num, value in enumerate(summary_df.columns.values):
            ws.write(0, col_num, value, header_fmt)
        ws.set_column("A:A", 16, cell_fmt)
        ws.set_column("B:B", 28, cell_fmt)
        ws.set_column("C:F", 10, center_fmt)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(summary_df), 1), len(summary_df.columns) - 1)
        # Total 列加粗居中
        total_col = summary_df.columns.get_loc("Total")
        for row in range(1, len(summary_df) + 1):
            ws.write(row, total_col, int(summary_df.iloc[row - 1, total_col]), total_fmt)
        # 未识别/无库位高亮
        if len(summary_df) > 0:
            ws.conditional_format(1, 0, len(summary_df), len(summary_df.columns) - 1, {
                "type": "text", "criteria": "containing", "value": "未识别", "format": warn_fmt
            })
            ws.conditional_format(1, 0, len(summary_df), len(summary_df.columns) - 1, {
                "type": "text", "criteria": "containing", "value": "无库位", "format": warn_fmt
            })

        ws2 = writer.sheets["订单明细"]
        for col_num, value in enumerate(detail_df.columns.values):
            ws2.write(0, col_num, value, header_fmt)
        ws2.set_column("A:A", 14)
        ws2.set_column("B:B", 8)
        ws2.set_column("C:C", 26)
        ws2.set_column("D:E", 14)
        ws2.set_column("F:G", 10)
        ws2.set_column("H:I", 26)
        ws2.set_column("J:J", 14)
        ws2.freeze_panes(1, 0)
        ws2.autofilter(0, 0, max(len(detail_df), 1), len(detail_df.columns) - 1)

    output.seek(0)
    return output.getvalue()


# =========================
# Streamlit 页面
# =========================
st.set_page_config(page_title="独立站 Slip 检货单生成器", layout="wide")
st.title("独立站 Slip 检货单生成器")
st.caption("上传产品图册 + 独立站 Packing Slip PDF，自动生成：库位 / Product Name / S / M / L / Total。匹配规则：SKU 优先，款式名兜底。")

with st.sidebar:
    st.header("上传文件")
    catalog_file = st.file_uploader("1）产品图册 CSV / Excel", type=["csv", "xlsx", "xls"])
    slip_file = st.file_uploader("2）独立站 Packing Slip PDF", type=["pdf"])
    st.info("Labels PDF 只有 tracking/order 信息，生成检货单时通常不需要上传。匹配规则：先按 SKU 匹配；SKU 找不到时，再用款式名兜底。")

if catalog_file and slip_file:
    try:
        catalog_df = read_uploaded_table(catalog_file)
        sku_map, name_map = build_catalog_maps(catalog_df)
        details_df = parse_slip_pdf(slip_file)
        summary_df, enriched_df = enrich_and_summarize(details_df, sku_map, name_map)

        st.success(f"生成成功：共识别 {int(enriched_df['Qty'].sum())} 件商品，{len(summary_df)} 个检货款式。")

        col1, col2, col3 = st.columns(3)
        col1.metric("总件数", int(summary_df["Total"].sum()) if not summary_df.empty else 0)
        col2.metric("款式行数", len(summary_df))
        col3.metric("未识别/无库位行", int(summary_df["库位"].astype(str).str.contains("未识别|无库位", regex=True).sum()) if not summary_df.empty else 0)

        st.subheader("检货单预览")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        with st.expander("查看订单明细 / 匹配方式"):
            st.dataframe(enriched_df, use_container_width=True, hide_index=True)

        excel_bytes = make_excel(summary_df, enriched_df)
        csv_bytes = summary_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        st.download_button(
            label="下载 Excel 检货单",
            data=excel_bytes,
            file_name="独立站检货单.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            label="下载 CSV 检货单",
            data=csv_bytes,
            file_name="独立站检货单.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"生成失败：{e}")
        st.caption("建议检查：产品图册是否包含 SKU / 款式英文名称 / 库位；上传的 PDF 是否为 Packing Slip，不是 Label。")
else:
    st.warning("请先上传产品图册和独立站 Packing Slip PDF。")
