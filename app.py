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


def enrich_and_summarize(details_df: pd.DataFrame, sku_map: Dict[str, dict], name_map: Dict[str, dict]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    匹配规则：
    1）检货单生成永远 SKU 优先；
    2）SKU 找不到时，才用款式名兜底；
    3）如果 Slip 款式名和 Slip SKU 在图册里的对应关系不一致，单独输出 SKU 异常明细。
    """
    enriched = details_df.copy()
    product_names = []
    locations = []
    match_types = []
    sku_catalog_names = []
    name_catalog_skus = []
    issue_notes = []
    mismatch_rows = []

    for _, r in enriched.iterrows():
        order_no = clean_cell(r.get("Order #", ""))
        base_sku = clean_cell(r.get("Base SKU", "")).upper()
        full_sku = clean_cell(r.get("Full SKU", "")).upper()
        slip_name = clean_cell(r.get("Slip Product Name", ""))
        size = clean_cell(r.get("Size", "")).upper()
        name_key = norm_text(slip_name)

        sku_match = sku_map.get(base_sku) if base_sku else None
        name_match = name_map.get(name_key) if name_key else None

        sku_catalog_name = clean_cell(sku_match.get("product_name", "")) if sku_match else ""
        name_catalog_sku_raw = clean_cell(name_match.get("catalog_sku", "")) if name_match else ""
        name_catalog_base_sku = re.sub(r"-(S|M|L)$", "", name_catalog_sku_raw.strip(), flags=re.I).upper() if name_catalog_sku_raw else ""

        match = None
        match_type = "未匹配"
        issue_note = ""

        # 第一优先级：按 slip 里的 Base SKU 匹配产品图册。
        # 仓库实际检货以 SKU 对应的图册款式和库位为准。
        if sku_match:
            match = sku_match
            names_different = bool(slip_name and sku_catalog_name and norm_text(sku_catalog_name) != name_key)
            # 只有当 Slip 款式名能在图册中找到，且该款式对应的 SKU 与 Slip SKU 不同时，才判定为真正 SKU 不符。
            # 例如：Slip 写 Citrus Veil / NOF031-L，但图册里 Citrus Veil = NOF030，而 NOF031 = Lady Cherry。
            if names_different and name_catalog_base_sku and name_catalog_base_sku != base_sku:
                match_type = "SKU匹配(名称不一致)"
                issue_note = (
                    f"Slip上【{slip_name}】的SKU是【{full_sku or base_sku}】，"
                    f"但产品图册中【{slip_name}】对应SKU是【{name_catalog_base_sku}】；"
                    f"同时【{base_sku}】在图册中对应【{sku_catalog_name}】。"
                )

                mismatch_rows.append({
                    "Order": order_no,
                    "Slip 款式名": slip_name,
                    "Slip SKU": full_sku or base_sku,
                    "Slip Base SKU": base_sku,
                    "图册中同款式对应 SKU": name_catalog_base_sku,
                    "SKU 在图册中对应的款式": sku_catalog_name or "图册未找到该SKU",
                    "检货实际采用": "按 Slip SKU 对应的图册款式/库位",
                    "异常说明": "Slip 款式名和 Slip SKU 在产品图册里的对应关系不一致",
                })
            elif names_different and not (name_key in norm_text(sku_catalog_name) or norm_text(sku_catalog_name) in name_key):
                match_type = "SKU匹配(名称需检查)"
                issue_note = (
                    f"Slip款式名【{slip_name}】和图册中该SKU对应款式【{sku_catalog_name}】不完全一致，"
                    "但未在图册中找到同名款式，暂不判定为SKU不符。"
                )
            else:
                match_type = "SKU匹配"
        elif name_match:
            match = name_match
            match_type = "款式名兜底匹配(SKU未找到)"
            issue_note = (
                f"Slip SKU【{full_sku or base_sku}】在产品图册中未找到，"
                f"已用款式名【{slip_name}】兜底匹配。"
            )
            mismatch_rows.append({
                "Order": order_no,
                "Slip 款式名": slip_name,
                "Slip SKU": full_sku or base_sku,
                "Slip Base SKU": base_sku,
                "图册中同款式对应 SKU": name_catalog_base_sku or "图册未找到该款式名",
                "SKU 在图册中对应的款式": "图册未找到该SKU",
                "检货实际采用": "按款式名兜底匹配",
                "异常说明": "Slip SKU 在产品图册中未找到",
            })

        if match:
            product_name = match.get("product_name") or slip_name
            location = match.get("location") or ""
        else:
            product_name = slip_name or base_sku
            location = ""
            issue_note = issue_note or "Slip SKU 和款式名都没有在产品图册中匹配到。"

        if not location:
            location = "无库位(特殊款)" if is_special_item(product_name, base_sku, size) else "未识别库位"

        product_names.append(product_name)
        locations.append(location)
        match_types.append(match_type)
        sku_catalog_names.append(sku_catalog_name)
        name_catalog_skus.append(name_catalog_base_sku)
        issue_notes.append(issue_note)

    enriched["Product Name"] = product_names
    enriched["库位"] = locations
    enriched["匹配方式"] = match_types
    enriched["SKU在图册中对应款式"] = sku_catalog_names
    enriched["图册中同款式对应SKU"] = name_catalog_skus
    enriched["异常说明"] = issue_notes

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

    mismatch_df = pd.DataFrame(mismatch_rows)
    if mismatch_df.empty:
        mismatch_df = pd.DataFrame(columns=[
            "Order", "Slip 款式名", "Slip SKU", "Slip Base SKU", "图册中同款式对应 SKU",
            "SKU 在图册中对应的款式", "检货实际采用", "异常说明"
        ])

    summary = summary.sort_values(by="库位", key=lambda col: col.map(location_sort_key)).reset_index(drop=True)
    return summary, enriched, mismatch_df

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


def make_excel(summary_df: pd.DataFrame, detail_df: pd.DataFrame, mismatch_df: pd.DataFrame) -> bytes:
    """生成带格式的 Excel。"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="检货单")
        detail_df.to_excel(writer, index=False, sheet_name="订单明细")
        mismatch_df.to_excel(writer, index=False, sheet_name="SKU异常明细")

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

        ws3 = writer.sheets["SKU异常明细"]
        for col_num, value in enumerate(mismatch_df.columns.values):
            ws3.write(0, col_num, value, header_fmt)
        ws3.set_column("A:A", 14)
        ws3.set_column("B:B", 26)
        ws3.set_column("C:D", 16)
        ws3.set_column("E:F", 30)
        ws3.set_column("G:G", 28)
        ws3.set_column("H:H", 42)
        ws3.freeze_panes(1, 0)
        ws3.autofilter(0, 0, max(len(mismatch_df), 1), len(mismatch_df.columns) - 1)

    output.seek(0)
    return output.getvalue()


# =========================
# Streamlit 页面
# =========================
st.set_page_config(page_title="独立站 Slip 检货单生成器", layout="wide")
st.title("独立站 Slip 检货单生成器")
st.caption("上传产品图册 + 独立站 Packing Slip PDF，自动生成：库位 / Product Name / S / M / L / Total。匹配规则：SKU 优先，款式名兜底；若 SKU 和款式名对应关系不一致，会单独显示 SKU 异常明细。")

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
        summary_df, enriched_df, mismatch_df = enrich_and_summarize(details_df, sku_map, name_map)

        st.success(f"生成成功：共识别 {int(enriched_df['Qty'].sum())} 件商品，{len(summary_df)} 个检货款式。")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("总件数", int(summary_df["Total"].sum()) if not summary_df.empty else 0)
        col2.metric("款式行数", len(summary_df))
        col3.metric("未识别/无库位行", int(summary_df["库位"].astype(str).str.contains("未识别|无库位", regex=True).sum()) if not summary_df.empty else 0)
        col4.metric("SKU异常", len(mismatch_df))

        if not mismatch_df.empty:
            st.warning("发现 SKU / 款式名对应关系异常：检货单仍然优先按照 Slip SKU 对应的产品图册款式和库位生成。")
            st.subheader("SKU 异常明细")
            st.dataframe(mismatch_df, use_container_width=True, hide_index=True)

        st.subheader("检货单预览")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        with st.expander("查看订单明细 / 匹配方式"):
            st.dataframe(enriched_df, use_container_width=True, hide_index=True)

        excel_bytes = make_excel(summary_df, enriched_df, mismatch_df)
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
