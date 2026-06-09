# -*- coding: utf-8 -*-
"""
NailVesta 出單達人 vs 深達名單 比對工具
====================================
針對實際檔案格式打造:
  - 出單達人名單 = TikTok「Creator List」匯出檔
        辨識欄位: Creator username
        出單判定: Affiliate orders > 0(可切換成 GMV / Items sold)
  - 深達名單     = NailVesta_深度达人List
        辨識欄位: handle(工作表「深度达人List」)

比對邏輯:
  出單達人(4 月 ∪ 5 月,可多檔)中,handle 不在深達名單者。
  比對前統一去掉 @、轉小寫、清空白。

執行:
  pip install streamlit pandas openpyxl
  streamlit run 深達名單比對.py
"""

import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="出單達人 vs 深達名單 比對", layout="wide")


def to_num(series: pd.Series) -> pd.Series:
    """把含逗號 / -- 的字串欄位轉成數字,無法轉的當 0。"""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "").str.replace("--", "").str.strip(),
        errors="coerce",
    ).fillna(0)


def normalize(value) -> str:
    """達人 handle 標準化:去頭尾空白 → 去開頭 @ → 壓縮空白 → 轉小寫。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip().lstrip("@").strip()
    return re.sub(r"\s+", " ", s).lower()


def pick_column(columns, candidates, fallback_keywords):
    """先找完全相符的候選欄名,再用關鍵字猜,最後退回第一欄。"""
    cols = list(columns)
    for cand in candidates:
        if cand in cols:
            return cand
    for col in cols:
        low = str(col).lower()
        if any(k in low for k in fallback_keywords):
            return col
    return cols[0] if cols else None


def read_creator_list(uploaded_file):
    """讀取單一 Creator List。回傳 (df, 預設達人欄)。"""
    df = pd.read_excel(uploaded_file, dtype=str).fillna("")
    creator_col = pick_column(
        df.columns,
        candidates=["Creator username"],
        fallback_keywords=["creator", "username", "handle", "达人", "達人"],
    )
    return df, creator_col


st.title("📊 出單達人 vs 深達名單 比對工具")
st.caption("找出「有出單,但不在深達名單」的達人。比對前自動處理 @、大小寫、空白差異。")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1️⃣ 出單達人名單(Creator List)")
    st.write("上傳 4 月、5 月的 Creator List,可一次多檔。")
    order_files = st.file_uploader(
        "Creator List(可多檔)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="orders",
    )
with c2:
    st.subheader("2️⃣ 深達名單")
    st.write("上傳 NailVesta 深度达人 List。")
    deep_files = st.file_uploader(
        "深達名單(可多檔)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="deep",
    )

if not order_files or not deep_files:
    st.info("👆 請先上傳「出單達人名單」與「深達名單」,才會開始比對。")
    st.stop()

st.divider()
st.subheader("⚙️ 出單判定設定")
metric_label = st.radio(
    "「出單」的判定依據(該欄位 > 0 才算有出單)",
    options=["Affiliate orders(訂單數)", "Affiliate GMV(銷售額)", "Items sold(售出件數)"],
    index=0,
    horizontal=True,
)
metric_map = {
    "Affiliate orders(訂單數)": "Affiliate orders",
    "Affiliate GMV(銷售額)": "Affiliate GMV",
    "Items sold(售出件數)": "Items sold",
}
metric_col = metric_map[metric_label]

st.divider()
st.subheader("📂 出單達人名單 — 欄位確認")

month_frames = []
with st.expander("展開確認各檔欄位", expanded=True):
    for f in order_files:
        df, creator_col = read_creator_list(f)
        if df.empty:
            st.warning(f"⚠️「{f.name}」是空的,已略過。")
            continue

        cols = st.columns([2, 2])
        with cols[0]:
            creator_col = st.selectbox(
                f"「{f.name}」達人欄位",
                options=list(df.columns),
                index=list(df.columns).index(creator_col) if creator_col in df.columns else 0,
                key=f"creator_{f.name}",
            )
        with cols[1]:
            m = re.search(r"(\d{4})(\d{2})\d{2}-", f.name)
            default_tag = f"{m.group(1)}-{m.group(2)}" if m else f.name
            month_tag = st.text_input(
                f"「{f.name}」月份標籤", value=default_tag, key=f"tag_{f.name}"
            )

        metric_use = metric_col if metric_col in df.columns else None
        if metric_use is None:
            st.warning(f"⚠️「{f.name}」找不到欄位「{metric_col}」,此檔將視為全部都有出單。")

        work = pd.DataFrame({
            "key": df[creator_col].map(normalize),
            "達人": df[creator_col].astype(str).str.strip(),
        })
        work["GMV"] = to_num(df["Affiliate GMV"]) if "Affiliate GMV" in df.columns else 0
        work["單數"] = to_num(df["Affiliate orders"]) if "Affiliate orders" in df.columns else 0
        if metric_use is not None:
            work = work[to_num(df[metric_use]) > 0]
        work = work[work["key"] != ""]
        work["月份"] = month_tag
        month_frames.append(work)

        st.caption(f"✅「{f.name}」讀到出單達人 {len(work)} 位")

if not month_frames:
    st.error("出單達人名單沒有讀到任何有效資料。")
    st.stop()

orders_all = pd.concat(month_frames, ignore_index=True)

st.divider()
st.subheader("📂 深達名單 — 欄位確認")

deep_keys = set()
with st.expander("展開確認各檔欄位", expanded=True):
    for f in deep_files:
        try:
            xl = pd.ExcelFile(f)
            sheet = st.selectbox(
                f"「{f.name}」工作表",
                options=xl.sheet_names,
                index=(xl.sheet_names.index("深度达人List")
                       if "深度达人List" in xl.sheet_names else 0),
                key=f"sheet_{f.name}",
            )
            ddf = pd.read_excel(f, sheet_name=sheet, dtype=str).fillna("")
        except Exception:
            ddf = pd.read_csv(f, dtype=str).fillna("")

        handle_col = pick_column(
            ddf.columns,
            candidates=["handle", "Handle"],
            fallback_keywords=["handle", "creator", "username", "达人", "達人"],
        )
        handle_col = st.selectbox(
            f"「{f.name}」handle 欄位",
            options=list(ddf.columns),
            index=list(ddf.columns).index(handle_col) if handle_col in ddf.columns else 0,
            key=f"handle_{f.name}",
        )
        keys = set(ddf[handle_col].map(normalize)) - {""}
        deep_keys |= keys
        st.caption(f"✅「{f.name}」讀到深達 handle {len(keys)} 個(去重)")

if not deep_keys:
    st.error("深達名單沒有讀到任何 handle。")
    st.stop()

agg = (
    orders_all.groupby("key")
    .agg(
        达人=("達人", "first"),
        合計GMV=("GMV", lambda s: round(s.sum(), 2)),
        合計單數=("單數", lambda s: int(s.sum())),
        出現月份=("月份", lambda s: "+".join(sorted(set(s)))),
    )
    .reset_index()
)

total_unique = len(agg)
in_deep = agg["key"].isin(deep_keys).sum()

not_in_deep = agg[~agg["key"].isin(deep_keys)].copy()
not_in_deep = not_in_deep.sort_values("合計GMV", ascending=False).reset_index(drop=True)
not_in_deep = not_in_deep[["达人", "出現月份", "合計GMV", "合計單數"]]
not_in_deep.insert(0, "序号", range(1, len(not_in_deep) + 1))

st.divider()
st.subheader("✅ 比對結果")

m1, m2, m3 = st.columns(3)
m1.metric("出單達人總數(去重)", total_unique)
m2.metric("在深達名單", int(in_deep))
m3.metric("⚠️ 有出單但不在深達名單", len(not_in_deep))

st.markdown("#### 有出單但不在深達名單的達人(依合計 GMV 由高到低)")
if len(not_in_deep) == 0:
    st.success("🎉 所有出單達人都在深達名單內。")
else:
    st.dataframe(not_in_deep, use_container_width=True, hide_index=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        not_in_deep.to_excel(writer, index=False, sheet_name="不在深達名單")
    buffer.seek(0)
    st.download_button(
        "⬇️ 下載結果 (Excel)",
        data=buffer,
        file_name="出單但不在深達名單.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with st.expander("🔍 完整明細(所有出單達人 + 是否在深達名單)"):
    full = agg.copy()
    full["在深達名單"] = full["key"].isin(deep_keys).map({True: "是", False: "否"})
    full = full.drop(columns=["key"]).sort_values(["在深達名單", "合計GMV"], ascending=[True, False])
    st.dataframe(full, use_container_width=True, hide_index=True)
