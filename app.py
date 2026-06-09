# -*- coding: utf-8 -*-
"""
NailVesta 達人分析工具(多分頁)
================================
共用上傳:4/5 月 Creator List + 深度达人 List

分頁:
  ① 招募名單   — 出單但不在深達名單,並分級排優先序
  ② 深達健康度 — 啟動率 / 休眠達人(掛名單但 0 單)
  ③ 流失/月變化 — 流失、新增、持續,深達流失優先追
  ④ 效率與質量 — 退款率、成本率、AOV、轉化、粉絲產出

出單判定:Affiliate orders > 0(可切換)。
比對鍵:Creator username vs handle,統一去 @、轉小寫、清空白。

執行:
  pip install streamlit pandas openpyxl
  streamlit run app.py
"""

import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="NailVesta 達人分析", layout="wide")


# ---------------------------------------------------------------- 工具
def to_num(series: pd.Series) -> pd.Series:
    """含逗號 / -- / % 的字串轉數字,無法轉的當 0。"""
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("--", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip(),
        errors="coerce",
    ).fillna(0)


def normalize(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip().lstrip("@").strip()
    return re.sub(r"\s+", " ", s).lower()


def pick_column(columns, candidates, fallback_keywords):
    cols = list(columns)
    for cand in candidates:
        if cand in cols:
            return cand
    for col in cols:
        if any(k in str(col).lower() for k in fallback_keywords):
            return col
    return cols[0] if cols else None


def safe_ratio(num, den):
    """num/den*100,den=0 時回 NaN,並四捨五入到 1 位。"""
    den = den.replace(0, np.nan)
    return (num / den * 100).round(1)


def to_excel_bytes(df, sheet="Sheet1"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    return buf


METRIC_COLS = {
    "GMV": "Affiliate GMV",
    "orders": "Affiliate orders",
    "items": "Items sold",
    "comm": "Est. commission",
    "flat": "Est. flat fee",
    "refGMV": "Affiliate refunded GMV",
    "refItems": "Affiliate items refunded",
    "impr": "Product impressions",
    "ctr": "CTR",
    "followers": "Affiliate followers",
    "aov": "Avg. order value",
    "liveGMV": "Affiliate LIVE GMV",
    "videoGMV": "Affiliate shoppable video GMV",
    "showcaseGMV": "Affiliate showcase GMV",
}


# ---------------------------------------------------------------- 上傳
st.title("📊 NailVesta 達人分析工具")
st.caption("一次上傳 4/5 月 Creator List 與深達名單,下方分頁提供多種分析。")

c1, c2 = st.columns(2)
with c1:
    st.subheader("1️⃣ 出單達人名單(Creator List)")
    order_files = st.file_uploader(
        "Creator List(可多檔)", type=["xlsx", "xls", "csv"],
        accept_multiple_files=True, key="orders",
    )
with c2:
    st.subheader("2️⃣ 深達名單")
    deep_files = st.file_uploader(
        "深達名單(可多檔)", type=["xlsx", "xls", "csv"],
        accept_multiple_files=True, key="deep",
    )

if not order_files or not deep_files:
    st.info("👆 請先上傳「出單達人名單」與「深達名單」。")
    st.stop()

metric_label = st.radio(
    "「出單」判定依據(該欄 > 0 才算出單)",
    ["Affiliate orders(訂單數)", "Affiliate GMV(銷售額)", "Items sold(售出件數)"],
    index=0, horizontal=True,
)
metric_key = {"Affiliate orders(訂單數)": "orders",
              "Affiliate GMV(銷售額)": "GMV",
              "Items sold(售出件數)": "items"}[metric_label]


# ---------------------------------------------------------------- 讀檔
@st.cache_data(show_spinner=False)
def parse_order_file(name, data, metric_key):
    df = pd.read_excel(io.BytesIO(data), dtype=str).fillna("") if not name.lower().endswith(".csv") \
        else pd.read_csv(io.BytesIO(data), dtype=str).fillna("")
    creator_col = pick_column(df.columns, ["Creator username"],
                              ["creator", "username", "handle", "达人", "達人"])
    out = pd.DataFrame({"key": df[creator_col].map(normalize),
                        "達人": df[creator_col].astype(str).str.strip()})
    for k, col in METRIC_COLS.items():
        out[k] = to_num(df[col]) if col in df.columns else 0
    out = out[(out["key"] != "") & (out[metric_key] > 0)].reset_index(drop=True)
    return out


def month_tag(name):
    m = re.search(r"(\d{4})(\d{2})\d{2}-", name)
    return f"{m.group(1)}-{m.group(2)}" if m else name


frames = []
for f in order_files:
    part = parse_order_file(f.name, f.getvalue(), metric_key)
    part["月份"] = month_tag(f.name)
    frames.append(part)
orders_all = pd.concat(frames, ignore_index=True)
months = sorted(orders_all["月份"].unique())


@st.cache_data(show_spinner=False)
def parse_deep_file(name, data):
    try:
        xl = pd.ExcelFile(io.BytesIO(data))
        sheet = "深度达人List" if "深度达人List" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, dtype=str).fillna("")
    except Exception:
        df = pd.read_csv(io.BytesIO(data), dtype=str).fillna("")
    hcol = pick_column(df.columns, ["handle", "Handle"],
                       ["handle", "creator", "username", "达人", "達人"])
    df["key"] = df[hcol].map(normalize)
    df = df[df["key"] != ""].copy()
    keep = ["key"]
    for c in ["handle", "Level", "评级", "深度前出单数", "深度合作Status"]:
        if c in df.columns:
            keep.append(c)
    return df[keep].drop_duplicates("key", keep="first")


deep_parts = [parse_deep_file(f.name, f.getvalue()) for f in deep_files]
deep = pd.concat(deep_parts, ignore_index=True).drop_duplicates("key", keep="first")
deep_keys = set(deep["key"])
order_keys = set(orders_all["key"])

# 每位達人跨月彙總
agg = (orders_all.groupby("key")
       .agg(達人=("達人", "first"),
            GMV=("GMV", "sum"), orders=("orders", "sum"), items=("items", "sum"),
            comm=("comm", "sum"), flat=("flat", "sum"),
            refGMV=("refGMV", "sum"), refItems=("refItems", "sum"),
            impr=("impr", "sum"), followers=("followers", "max"),
            ctr=("ctr", "mean"),
            liveGMV=("liveGMV", "sum"), videoGMV=("videoGMV", "sum"),
            showcaseGMV=("showcaseGMV", "sum"),
            出現月份=("月份", lambda s: "+".join(sorted(set(s)))))
       .reset_index())
agg["在深達名單"] = agg["key"].isin(deep_keys)

# 各月單數寬表
pivot = orders_all.pivot_table(index="key", columns="月份", values="orders",
                               aggfunc="sum", fill_value=0)


# ================================================================ 分頁
tab1, tab2, tab3, tab4 = st.tabs(
    ["① 招募名單", "② 深達健康度", "③ 流失/月變化", "④ 效率與質量"])

# ---------------- ① 招募名單 ----------------
with tab1:
    st.subheader("出單但不在深達名單 — 招募優先序")
    nd = agg[~agg["在深達名單"]].copy()

    def classify(k):
        vals = {m: pivot.loc[k, m] if (m in pivot.columns and k in pivot.index) else 0
                for m in months}
        nonzero = [m for m in months if vals[m] > 0]
        if len(nonzero) >= 2:
            return "A. 連續多月(穩定)"
        if len(months) >= 2 and vals.get(months[-1], 0) > vals.get(months[0], 0):
            return "B. 上升中"
        return "C. 單月出單"

    nd["分級"] = nd["key"].map(classify)
    nd = nd.sort_values(["分級", "GMV"], ascending=[True, False]).reset_index(drop=True)
    show = nd[["達人", "分級", "出現月份", "GMV", "orders", "items", "followers"]].copy()
    show.columns = ["達人", "分級", "出現月份", "合計GMV", "合計單數", "合計件數", "粉絲數"]
    show.insert(0, "序号", range(1, len(show) + 1))

    a, b, c = st.columns(3)
    a.metric("不在深達(可招募)", len(nd))
    b.metric("其中連續多月", int((nd["分級"].str.startswith("A")).sum()))
    c.metric("其中上升中", int((nd["分級"].str.startswith("B")).sum()))
    st.caption("分級:A 連續多月都出單(最該優先簽)> B 出單量月增(成長中)> C 只單月出單。")
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.download_button("⬇️ 下載招募名單", to_excel_bytes(show, "招募名單"),
                       "招募名單_不在深達.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------- ② 深達健康度 ----------------
with tab2:
    st.subheader("深達名單健康度")
    activated = deep_keys & order_keys
    dormant = deep_keys - order_keys
    rate = len(activated) / len(deep_keys) * 100 if deep_keys else 0

    a, b, c, d = st.columns(4)
    a.metric("深達名單總數", len(deep_keys))
    b.metric("有出單(已啟動)", len(activated))
    c.metric("啟動率", f"{rate:.1f}%")
    d.metric("休眠(0 單)", len(dormant))

    st.markdown("#### 💤 休眠深達達人(掛名單但這段期間 0 單)")
    dorm = deep[deep["key"].isin(dormant)].copy()
    dorm = dorm.drop(columns=["key"])
    st.caption(f"共 {len(dorm)} 位。可考慮喚醒或從名單汰除。")
    st.dataframe(dorm, use_container_width=True, hide_index=True)
    st.download_button("⬇️ 下載休眠名單", to_excel_bytes(dorm, "休眠深達"),
                       "休眠深達名單.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------- ③ 流失/月變化 ----------------
with tab3:
    st.subheader("流失 / 月變化")
    if len(months) < 2:
        st.info("只上傳了一個月的資料,無法做月變化比較。請同時上傳兩個月的 Creator List。")
    else:
        m0, m1 = months[0], months[-1]
        s0 = set(orders_all[orders_all["月份"] == m0]["key"])
        s1 = set(orders_all[orders_all["月份"] == m1]["key"])
        churned, newly, kept = s0 - s1, s1 - s0, s0 & s1
        churned_deep = churned & deep_keys

        a, b, c, d = st.columns(4)
        a.metric(f"{m0} 出單", len(s0))
        b.metric(f"{m1} 出單", len(s1))
        c.metric("流失(前有後無)", len(churned))
        d.metric("新增(前無後有)", len(newly))
        st.warning(f"⚠️ 流失達人中有 **{len(churned_deep)}** 位是深達達人,建議優先回訪。")

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown(f"#### 🔻 流失達人({m0} 有,{m1} 無)")
            ch = agg[agg["key"].isin(churned)][["達人", "GMV", "orders"]].copy()
            ch["是否深達"] = ch.index.map(lambda i: "是" if agg.loc[i, "key"] in deep_keys else "否")
            ch = ch.sort_values("GMV", ascending=False)
            ch.columns = ["達人", f"{m0}+期間GMV", "期間單數", "是否深達"]
            st.dataframe(ch, use_container_width=True, hide_index=True)
        with col_r:
            st.markdown(f"#### 🆕 新增達人({m1} 才出單)")
            nw = agg[agg["key"].isin(newly)][["達人", "GMV", "orders"]].copy()
            nw["在深達"] = nw["達人"].map(lambda x: "")  # placeholder
            nw = agg[agg["key"].isin(newly)][["達人", "GMV", "orders", "在深達名單"]].copy()
            nw["在深達名單"] = nw["在深達名單"].map({True: "是", False: "否"})
            nw = nw.sort_values("GMV", ascending=False)
            nw.columns = ["達人", "期間GMV", "期間單數", "在深達"]
            st.dataframe(nw, use_container_width=True, hide_index=True)

        st.markdown(f"#### 🔁 持續出單達人 GMV 變化({m0} → {m1})")
        g0 = orders_all[orders_all["月份"] == m0].groupby("key")["GMV"].sum()
        g1 = orders_all[orders_all["月份"] == m1].groupby("key")["GMV"].sum()
        kp = pd.DataFrame({"key": list(kept)})
        kp["達人"] = kp["key"].map(agg.set_index("key")["達人"])
        kp[f"{m0} GMV"] = kp["key"].map(g0).round(2)
        kp[f"{m1} GMV"] = kp["key"].map(g1).round(2)
        kp["變化"] = (kp[f"{m1} GMV"] - kp[f"{m0} GMV"]).round(2)
        kp["趨勢"] = np.where(kp["變化"] > 0, "↑ 成長", np.where(kp["變化"] < 0, "↓ 下滑", "—"))
        kp = kp.drop(columns=["key"]).sort_values("變化", ascending=False)
        st.dataframe(kp, use_container_width=True, hide_index=True)

# ---------------- ④ 效率與質量 ----------------
with tab4:
    st.subheader("效率與質量")
    e = agg.copy()
    e["退款率%"] = safe_ratio(e["refGMV"], e["GMV"])
    e["成本率%"] = safe_ratio(e["comm"] + e["flat"], e["GMV"])
    e["AOV"] = (e["GMV"] / e["orders"].replace(0, np.nan)).round(2)
    e["每千次曝光出單"] = (e["orders"] / e["impr"].replace(0, np.nan) * 1000).round(2)
    e["GMV"] = e["GMV"].round(2)

    tot_ref = e["refGMV"].sum() / e["GMV"].sum() * 100 if e["GMV"].sum() else 0
    tot_cost = (e["comm"].sum() + e["flat"].sum()) / e["GMV"].sum() * 100 if e["GMV"].sum() else 0
    a, b, c = st.columns(3)
    a.metric("整體退款率", f"{tot_ref:.1f}%")
    a.caption("退款 GMV ÷ 總 GMV")
    b.metric("整體成本率", f"{tot_cost:.1f}%")
    b.caption("(佣金+固定費) ÷ 總 GMV")
    c.metric("出單達人數", len(e))

    st.markdown("#### ⚠️ 高退款率達人(GMV ≥ 50,退款率前 20)")
    hi_ref = e[e["GMV"] >= 50].sort_values("退款率%", ascending=False).head(20)
    st.dataframe(hi_ref[["達人", "GMV", "退款率%", "refItems"]]
                 .rename(columns={"refItems": "退款件數"}),
                 use_container_width=True, hide_index=True)

    st.markdown("#### 💰 高成本率達人(GMV ≥ 50,成本率前 20)")
    hi_cost = e[e["GMV"] >= 50].sort_values("成本率%", ascending=False).head(20)
    st.dataframe(hi_cost[["達人", "GMV", "成本率%", "comm", "flat"]]
                 .rename(columns={"comm": "佣金", "flat": "固定費"}),
                 use_container_width=True, hide_index=True)

    st.markdown("#### 🌱 高粉絲但低產出(粉絲前 30% 中 GMV 最低,潛力未開發)")
    if len(e) > 0:
        thr = e["followers"].quantile(0.7)
        pot = e[e["followers"] >= thr].sort_values("GMV").head(20)
        st.dataframe(pot[["達人", "followers", "GMV", "orders"]]
                     .rename(columns={"followers": "粉絲數", "orders": "單數"}),
                     use_container_width=True, hide_index=True)

    st.markdown("#### 🎬 內容形式 GMV 拆解(直播 / 短影片 / 櫥窗,前 20)")
    cont = e.sort_values("GMV", ascending=False).head(20)[
        ["達人", "GMV", "liveGMV", "videoGMV", "showcaseGMV"]].copy()
    cont.columns = ["達人", "總GMV", "直播GMV", "短影片GMV", "櫥窗GMV"]
    cont[["直播GMV", "短影片GMV", "櫥窗GMV"]] = cont[["直播GMV", "短影片GMV", "櫥窗GMV"]].round(2)
    st.dataframe(cont, use_container_width=True, hide_index=True)

    st.download_button("⬇️ 下載完整效率表", to_excel_bytes(
        e[["達人", "出現月份", "GMV", "orders", "items", "AOV",
           "退款率%", "成本率%", "每千次曝光出單", "followers", "在深達名單"]], "效率"),
        "達人效率明細.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
