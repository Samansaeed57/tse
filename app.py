import datetime
import io
import os
import finpy_tse as fpy
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# تنظیمات صفحه با Layout وسیع
st.set_page_config(
    page_title="داشبورد جامع بازار سرمایه و مانیتورینگ حرفه‌ای", layout="wide"
)

# تزریق CSS اختصاصی استریم‌لیت همراه با استایل رنگی سبز و قرمز برای مقادیر مثبت و منفی
st.markdown(
    """
    <style>
    @import url('https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css');
    
    html, body, [class*="css"], * {
        font-family: 'Vazirmatn', Tahoma, sans-serif !important;
        direction: rtl;
    }
    
    .stApp {
        background-color: #0e1525;
        color: #f8fafc;
    }
    
    .metric-box {
        background-color: #162032;
        border: 1px solid #1e293b;
        padding: 16px;
        border-radius: 14px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
        margin-bottom: 10px;
    }
    
    .signal-box {
        background-color: #162032;
        border-left: 5px solid #00e676;
        padding: 18px;
        border-radius: 10px;
        margin-top: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
    }
    .signal-box-warn {
        background-color: #162032;
        border-left: 5px solid #ff4d4d;
        padding: 18px;
        border-radius: 10px;
        margin-top: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
    }
    
    div.stButton > button {
        background-color: #00e676 !important;
        color: #0f172a !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
    }
    div.stButton > button:hover {
        background-color: #00c853 !important;
        color: #0f172a !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #162032;
        padding: 10px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #1e293b !important;
        color: #cbd5e1 !important;
        border-radius: 8px !important;
        padding: 8px 16px;
        font-weight: bold;
    }
    .stTabs [aria-selected="true"] {
        background-color: #00e676 !important;
        color: #0f172a !important;
    }
    
    div[data-testid="stDataFrame"] {
        direction: rtl;
    }
    </style>
    
    <meta http-equiv="refresh" content="300">
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=60)
def load_market_data():
  try:
    raw_data = fpy.Get_MarketWatch()
    if isinstance(raw_data, tuple):
      df = raw_data[0]
    else:
      df = raw_data
    return df
  except Exception as e:
    st.error(f"خطا در دریافت داده‌ها: {e}")
    return None


def convert_df_to_excel(df):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="MarketWatch", index=False)
  return output.getvalue()


def normalize_persian(text):
  if isinstance(text, str):
    return (
        text.replace("ي", "ی")
        .replace("ك", "ک")
        .replace("‌", "")
        .strip()
    )
  return text


def color_surplus_df(val):
  if isinstance(val, (int, float)):
    if val < 0:
      return "color: #ff4d4d; font-weight: bold;"
    elif val > 0:
      return "color: #00e676; font-weight: bold;"
  return ""


# هدر بالای صفحه
header_col1, header_col2, header_col3 = st.columns([3, 1, 1])
with header_col1:
  st.markdown(
      "### 📈 داشبورد جامع و مانیتورینگ حرفه‌ای بازار سرمایه ایران"
  )
  st.caption(
      f"آخرین بروزرسانی خودکار: {datetime.datetime.now().strftime('%H:%M:%S')} (هر"
      " ۵ دقیقه)"
  )
with header_col2:
  if st.button("🔄 به‌روزرسانی دستی"):
    st.cache_data.clear()
    st.rerun()
with header_col3:
  if st.button("🗑️ پاکسازی تاریخچه روز قبل"):
    if os.path.exists("market_history.csv"):
      os.remove("market_history.csv")
    st.success("تاریخچه پاک شد!")
    st.rerun()

# مدیریت پورتفوی کاربر از طریق سایدبار
st.sidebar.markdown("### 💼 مدیریت و تحلیل پورتفو")
portfolio_file = st.sidebar.file_uploader(
    "آپلود فایل اکسل پورتفو", type=["xlsx", "xls"]
)

with st.spinner("در حال دریافت و پردازش اطلاعات بازار..."):
  df_market = load_market_data()

if df_market is not None and not df_market.empty:
  if "Ticker" not in df_market.columns:
    df_market = df_market.reset_index()
    if "index" in df_market.columns:
      df_market = df_market.rename(columns={"index": "Ticker"})
    elif len(df_market.columns) > 0:
      df_market = df_market.rename(
          columns={df_market.columns[0]: "Ticker"}
      )

  for col_name in ["Group", "Sector", "GroupName", "SubSector", "گروه"]:
    if col_name in df_market.columns and col_name != "Sector":
      df_market["Sector"] = df_market[col_name]
      break

  if "Trade Type" in df_market.columns:
    df_market = df_market[
        df_market["Trade Type"].astype(str).str.strip().isin(["تابلو", "صندوق"])
    ]

  exclude_keywords = [
      "درآمد ثابت",
      "اختیار",
      "مرابحه",
      "سخاب",
      "اخزا",
      "اوراق بدهی",
      "صکوك",
      "صکوک",
      "سلف",
      "مشارکت",
      "اجاره",
  ]
  mask = pd.Series(True, index=df_market.index)
  for col in df_market.columns:
    if df_market[col].dtype == "object":
      for kw in exclude_keywords:
        mask = mask & (
            ~df_market[col]
            .astype(str)
            .apply(normalize_persian)
            .str.contains(normalize_persian(kw), na=False)
        )
  df_filtered = df_market[mask].copy()

  if "Ticker" in df_filtered.columns:
    df_filtered = df_filtered[
        ~df_filtered["Ticker"].astype(str).str.endswith("ح")
    ]
    df_filtered = df_filtered[
        ~df_filtered["Ticker"].astype(str).str.contains("ح-", na=False)
    ]

  numeric_cols = [
      "Vol_Buy_R",
      "No_Buy_R",
      "Vol_Sell_R",
      "No_Sell_R",
      "Close",
      "Final",
      "Volume",
      "Value",
      "Close(%)",
      "Final(%)",
      "PE",
  ]
  for col in numeric_cols:
    if col in df_filtered.columns:
      df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce")

  if all(
      c in df_filtered.columns
      for c in [
          "Vol_Buy_R",
          "No_Buy_R",
          "Vol_Sell_R",
          "No_Sell_R",
          "Close",
      ]
  ):
    df_filtered["Buy_Per_Capita"] = (
        df_filtered["Vol_Buy_R"] / df_filtered["No_Buy_R"]
    ).round(0)
    df_filtered["Sell_Per_Capita"] = (
        df_filtered["Vol_Sell_R"] / df_filtered["No_Sell_R"]
    ).round(0)
    df_filtered["Buyer_Power"] = (
        df_filtered["Buy_Per_Capita"]
        / df_filtered["Sell_Per_Capita"].replace(0, 1)
    ).round(2)
    df_filtered["Net_Real_Flow"] = (
        (df_filtered["Vol_Buy_R"] - df_filtered["Vol_Sell_R"])
        * df_filtered["Close"]
    ) / 10**10

  df_filtered["Value_Bln"] = (
      (df_filtered["Value"] / 10**10).round(1)
      if "Value" in df_filtered.columns
      else 0.0
  )

  bq_val_col = next(
      (
          c
          for c in ["BQ-Value", "BuyQueueValue", "BQ_Val"]
          if c in df_filtered.columns
      ),
      None,
  )
  sq_val_col = next(
      (
          c
          for c in ["SQ-Value", "SellQueueValue", "SQ_Val", "SellQueue"]
          if c in df_filtered.columns
      ),
      None,
  )

  df_filtered["Clean_BQ_Value"] = (
      pd.to_numeric(
          df_filtered[bq_val_col].astype(str).str.replace(",", "", regex=True),
          errors="coerce",
      ).fillna(0)
      / 10**10
  ) if bq_val_col else 0.0

  df_filtered["Clean_SQ_Value"] = (
      pd.to_numeric(
          df_filtered[sq_val_col].astype(str).str.replace(",", "", regex=True),
          errors="coerce",
      )
      .abs()
      .fillna(0)
      / 10**10
  ) if sq_val_col else 0.0

  total_symbols = len(df_filtered)
  pos_symbols = (
      len(df_filtered[df_filtered["Close(%)"] > 0])
      if "Close(%)" in df_filtered.columns
      else 0
  )
  neg_symbols = (
      len(df_filtered[df_filtered["Close(%)"] < 0])
      if "Close(%)" in df_filtered.columns
      else 0
  )
  neutral_symbols = total_symbols - (pos_symbols + neg_symbols)
  pos_pct = (
      int((pos_symbols / total_symbols) * 100) if total_symbols > 0 else 0
  )

  net_pos_neg = pos_symbols - neg_symbols
  market_turnover = pos_pct - 50
  net_queue = (
      df_filtered["Clean_BQ_Value"].sum() - df_filtered["Clean_SQ_Value"].sum()
  )
  mean_close_pct = (
      df_filtered["Close(%)"].mean() if "Close(%)" in df_filtered.columns else 0
  )
  excitement_osc = (
      (df_filtered["Close(%)"] - df_filtered["Final(%)"]).mean()
      if "Close(%)" in df_filtered.columns
      else 0
  )
  avg_buyer_power = (
      df_filtered["Buyer_Power"].mean()
      if "Buyer_Power" in df_filtered.columns
      else 1.0
  )
  net_total_flow = (
      df_filtered["Net_Real_Flow"].sum()
      if "Net_Real_Flow" in df_filtered.columns
      else 0
  )

  if "Close(%)" in df_filtered.columns:
    c_series = df_filtered["Close(%)"]
    count_bq = len(
        df_filtered[
            (df_filtered["Clean_BQ_Value"] > 0)
            | ((c_series >= 3.5) & (c_series <= 5.0))
        ]
    )
    count_gt3 = len(df_filtered[c_series > 3.0])
    count_1_3 = len(df_filtered[(c_series > 0) & (c_series <= 3.0)])
    count_neg1_1 = len(df_filtered[(c_series >= -1.0) & (c_series <= 0)])
    count_neg3_neg1 = len(df_filtered[(c_series >= -3.0) & (c_series < -1.0)])
    count_lt_neg3 = len(df_filtered[c_series < -3.0])
    count_sq = len(
        df_filtered[
            (df_filtered["Clean_SQ_Value"] > 0) | (c_series <= -4.5)
        ]
    )
  else:
    count_bq, count_gt3, count_1_3, count_neg1_1, count_neg3_neg1, count_lt_neg3, count_sq = (
        0,
        0,
        0,
        0,
        0,
        0,
        0,
    )

  max_dist_val = max(
      [
          total_symbols,
          1,
          count_bq,
          count_gt3,
          count_1_3,
          count_neg1_1,
          count_neg3_neg1,
          count_lt_neg3,
          count_sq,
      ]
  )

  current_time_str = datetime.datetime.now().strftime("%H:%M")
  current_date_str = datetime.datetime.now().strftime("%Y-%m-%d")
  history_file = "market_history.csv"
  date_file = "market_date.txt"

  if os.path.exists(date_file):
    with open(date_file, "r") as f:
      saved_date = f.read().strip()
    if saved_date != current_date_str:
      if os.path.exists(history_file):
        os.remove(history_file)
      with open(date_file, "w") as f:
        f.write(current_date_str)
  else:
    with open(date_file, "w") as f:
      f.write(current_date_str)

  if os.path.exists(history_file):
    df_history = pd.read_csv(history_file)
  else:
    df_history = pd.DataFrame(
        columns=[
            "زمان",
            "شاخص حال بازار",
            "چرخش بازار",
            "برآیند صفوف",
            "برآیند مثبت منفی",
            "اسیلاتور هیجان",
            "میانگین درصد تغییرات",
            "قدرت خریدار",
        ]
    )

  if df_history.empty or df_history.iloc[-1]["زمان"] != current_time_str:
    new_row = pd.DataFrame({
        "زمان": [current_time_str],
        "شاخص حال بازار": [pos_pct],
        "چرخش بازار": [market_turnover],
        "برآیند صفوف": [round(net_queue, 1)],
        "برآیند مثبت منفی": [net_pos_neg],
        "اسیلاتور هیجان": [round(excitement_osc, 2)],
        "میانگین درصد تغییرات": [round(mean_close_pct, 2)],
        "قدرت خریدار": [round(avg_buyer_power, 2)],
    })
    df_history = pd.concat([df_history, new_row], ignore_index=True)
    df_history.to_csv(history_file, index=False)

  # کارت‌های ۶ گانه بالای صفحه
  c1, c2, c3 = st.columns(3)
  with c1:
    st.markdown(
        f"""
        <div class="metric-box">
            <div style="color: #94a3b8; font-size: 14px;">تعادل مثبت / منفی</div>
            <div style="color: #00e676; font-size: 20px; font-weight: bold; margin-top: 5px;">مثبت {pos_symbols} / منفی {neg_symbols}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
  with c2:
    flow_color = "#00e676" if net_total_flow >= 0 else "#ff4d4d"
    st.markdown(
        f"""
        <div class="metric-box">
            <div style="color: #94a3b8; font-size: 14px;">خالص پول حقیقی</div>
            <div style="color: {flow_color}; font-size: 20px; font-weight: bold; margin-top: 5px;">{net_total_flow:+,.1f} ب.ت</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
  with c3:
    st.markdown(
        f"""
        <div class="metric-box">
            <div style="color: #94a3b8; font-size: 14px;">چرخش بازار</div>
            <div style="color: #38bdf8; font-size: 20px; font-weight: bold; margin-top: 5px;">{market_turnover:+d}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

  c4, c5, c6 = st.columns(3)
  with c4:
    st.markdown(
        f"""
        <div class="metric-box">
            <div style="color: #94a3b8; font-size: 14px;">هیجان بازار (میانگین)</div>
            <div style="color: #fbbf24; font-size: 20px; font-weight: bold; margin-top: 5px;">{mean_close_pct:+.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
  with c5:
    st.markdown(
        f"""
        <div class="metric-box">
            <div style="color: #94a3b8; font-size: 14px;">شاخص حال بازار</div>
            <div style="color: #00e676; font-size: 20px; font-weight: bold; margin-top: 5px;">{pos_pct}%</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
  with c6:
    phase = (
        "🟢 پر تقاضا"
        if pos_pct >= 55
        else ("🔴 پر عرضه" if pos_pct < 45 else "🟡 متعادل")
    )
    st.markdown(
        f"""
        <div class="metric-box">
            <div style="color: #94a3b8; font-size: 14px;">فاز فعلی بازار</div>
            <div style="color: #f8fafc; font-size: 20px; font-weight: bold; margin-top: 5px;">{phase}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

  st.markdown("---")

  # تب‌های اصلی داشبورد
  tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
      "📈 خلاصه عملکرد و نمای کلی",
      "🏭 تحلیل صنایع و گروه‌ها",
      "🔍 دیدبان پیشرفته و پول هوشمند",
      "📊 برترین صف‌های خرید و فروش",
      "🤖 استراتژیست هوشمند بازار",
      "💼 تحلیل هوشمند پورتفوی من",
  ])

  with tab1:
    if net_total_flow > 0 and pos_pct >= 50:
      st.markdown(
          f"""
            <div class="signal-box">
                <div style="font-weight: bold; font-size: 16px; color: #00e676; margin-bottom: 5px;">💡 تحلیل هوشمند و سیگنال بازار: وضعیت مثبت و ورود پول</div>
                <div style="color: #cbd5e1; font-size: 14px;">بازار در فاز <b>{phase}</b> قرار دارد. ورود خالص پول حقیقی معادل <b>{net_total_flow:+,.1f} میلیارد تومان</b> و شاخص حال بازار در سطح <b>{pos_pct}%</b> نشان‌دهنده برتری نسبی تقاضا است. صنایع پیشرو و نمادهای دارای قدرت خریدار بالا در دیدبان پایش شوند.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
    else:
      st.markdown(
          f"""
            <div class="signal-box-warn">
                <div style="font-weight: bold; font-size: 16px; color: #ff4d4d; margin-bottom: 5px;">⚠️ تحلیل هوشمند و هشدار بازار: غلبه عرضه یا خروج پول</div>
                <div style="color: #cbd5e1; font-size: 14px;">بازار در فاز <b>{phase}</b> است. با توجه به خروج پول حقیقی (<b>{net_total_flow:+,.1f} ب.ت</b>) و فشار فروش، توصیه می‌شود از خرید هیجانی پرهیز کرده و مدیریت ریسک و رصد صفوف فروش در اولویت قرار گیرد.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )

    ch_col1, ch_col2 = st.columns(2)

    with ch_col1:
      st.markdown("### 🍩 وسعت بازار (نمادهای مثبت، منفی و بدون تغییر)")
      breadth_df = pd.DataFrame({
          "وضعیت": ["مثبت", "منفی", "بدون تغییر"],
          "تعداد": [pos_symbols, neg_symbols, neutral_symbols],
      })
      fig_donut = px.pie(
          breadth_df,
          values="تعداد",
          names="وضعیت",
          hole=0.5,
          color="وضعیت",
          color_discrete_map={
              "مثبت": "#00e676",
              "منفی": "#ff4d4d",
              "بدون تغییر": "#90a4ae",
          },
      )
      fig_donut.update_layout(
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(color="#f8fafc", family="Vazirmatn", size=11),
          margin=dict(t=20, b=30, l=10, r=10),
          legend=dict(
              orientation="h",
              yanchor="top",
              y=-0.15,
              xanchor="center",
              x=0.5,
              font=dict(color="#ffffff", size=11),
          ),
      )
      st.plotly_chart(fig_donut, use_container_width=True)

      st.markdown("### 📊 اسیلاتور مشترک و چرخش کل بازار")
      fig_turnover = px.line(
          df_history,
          x="زمان",
          y=["چرخش بازار", "برآیند مثبت منفی"],
          markers=True,
      )
      fig_turnover.update_layout(
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(color="#ffffff", family="Vazirmatn", size=11),
          margin=dict(t=20, b=40, l=10, r=10),
          legend=dict(
              orientation="h",
              yanchor="top",
              y=-0.25,
              xanchor="center",
              x=0.5,
              font=dict(color="#ffffff", size=11),
              bgcolor="rgba(15, 23, 42, 0.9)",
              bordercolor="rgba(255,255,255,0.2)",
              borderwidth=1,
          ),
      )
      st.plotly_chart(fig_turnover, use_container_width=True)

      st.markdown("### 📉 روند درصد تغییرات قیمت سهام (هیجان)")
      fig_mean_pct = px.line(
          df_history,
          x="زمان",
          y="میانگین درصد تغییرات",
          markers=True,
      )
      fig_mean_pct.update_traces(line_color="#00e676", line_width=2.5)
      fig_mean_pct.update_layout(
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(color="#ffffff", family="Vazirmatn", size=11),
          margin=dict(t=20, b=20, l=10, r=10),
      )
      st.plotly_chart(fig_mean_pct, use_container_width=True)

    with ch_col2:
      st.markdown("### 📈 نمودار چندخطی رصد بازار و برآیند صفوف")
      fig_multi = go.Figure()
      fig_multi.add_trace(
          go.Scatter(
              x=df_history["زمان"],
              y=df_history["شاخص حال بازار"],
              mode="lines+markers",
              name="شاخص حال بازار (%)",
              line=dict(color="#00e676", width=3),
          )
      )
      fig_multi.add_trace(
          go.Scatter(
              x=df_history["زمان"],
              y=df_history["برآیند صفوف"],
              mode="lines",
              name="برآیند صفوف (همت)",
              line=dict(color="#38bdf8", width=2, dash="dash"),
          )
      )
      fig_multi.update_layout(
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(color="#ffffff", family="Vazirmatn", size=11),
          margin=dict(t=20, b=40, l=10, r=10),
          legend=dict(
              orientation="h",
              yanchor="top",
              y=-0.25,
              xanchor="center",
              x=0.5,
              font=dict(color="#ffffff", size=11),
              bgcolor="rgba(15, 23, 42, 0.9)",
              bordercolor="rgba(255,255,255,0.2)",
              borderwidth=1,
          ),
      )
      st.plotly_chart(fig_multi, use_container_width=True)

      st.markdown("### 🔮 روند قدرت خریدار")
      fig_bp = px.line(
          df_history,
          x="زمان",
          y="قدرت خریدار",
          markers=True,
      )
      fig_bp.update_traces(line_color="#e879f9", line_width=2.5)
      fig_bp.update_layout(
          paper_bgcolor="rgba(0,0,0,0)",
          plot_bgcolor="rgba(0,0,0,0)",
          font=dict(color="#ffffff", family="Vazirmatn", size=11),
          margin=dict(t=20, b=20, l=10, r=10),
      )
      st.plotly_chart(fig_bp, use_container_width=True)

    # بخش توزیع تغییر قیمت آخرین معامله (به شکل نوار افقی مطابق با تصویر ارسالی شما)
    st.markdown("---")
    st.markdown(
        f"### 📊 توزیع تغییر قیمت آخرین معامله &nbsp;&nbsp;|&nbsp;&nbsp; <span"
        f" style='font-size: 14px; color: #94a3b8;'>{total_symbols} نماد |"
        " باره‌ها: تعداد</span>",
        unsafe_allow_html=True,
    )

    dist_data = [
        ("صف خرید", count_bq, "#00c853"),
        ("بیش از مثبت سه درصد", count_gt3, "#00e676"),
        ("مثبت یک تا سه درصد", count_1_3, "#81c784"),
        ("منفی یک تا مثبت یک", count_neg1_1, "#90a4ae"),
        ("منفی سه تا منفی یک", count_neg3_neg1, "#e57373"),
        ("کمتر از منفی سه درصد", count_lt_neg3, "#e53935"),
        ("صف فروش", count_sq, "#d32f2f"),
    ]

    max_val = max([x[1] for x in dist_data]) if max([x[1] for x in dist_data]) > 0 else 1

    for label, val, color in dist_data:
      pct_width = int((val / max_val) * 100) if max_val > 0 else 0
      col_bar, col_val, col_lbl = st.columns([5, 1, 2])
      with col_bar:
        st.markdown(
            f"""
                <div style="background-color: #1e293b; border-radius: 4px; width: 100%; height: 18px; margin-top: 8px; overflow: hidden;">
                    <div style="background-color: {color}; width: {pct_width}%; height: 100%; border-radius: 4px;"></div>
                </div>
                """,
            unsafe_allow_html=True,
        )
      with col_val:
        st.markdown(
            f"<div style='color: {color}; font-weight: bold; text-align:"
            f" center;'>{val}</div>",
            unsafe_allow_html=True,
        )
      with col_lbl:
        st.markdown(
            f"<div style='color: #cbd5e1; text-align: right;'>{label}</div>",
            unsafe_allow_html=True,
        )

    st.markdown("---")

  with tab2:
    st.subheader("🏭 وضعیت گروه‌ها و صنایع بازار")
    if "Sector" in df_filtered.columns:
      ind_df = (
          df_filtered.groupby("Sector")
          .agg({
              "Ticker": "count",
              "Value_Bln": "sum",
              "Net_Real_Flow": "sum"
              if "Net_Real_Flow" in df_filtered.columns
              else "count",
              "Buyer_Power": "mean"
              if "Buyer_Power" in df_filtered.columns
              else "count",
          })
          .reset_index()
      )
      ind_df.columns = [
          "صنعت",
          "تعداد نماد",
          "ارزش معاملات (میلیارد تومان)",
          "خالص پول حقیقی (میلیارد تومان)",
          "میانگین قدرت خریدار",
      ]
      ind_df = ind_df.sort_values(
          by="ارزش معاملات (میلیارد تومان)", ascending=False
      )
      st.dataframe(
          ind_df.style.applymap(
              color_surplus_df,
              subset=[
                  "خالص پول حقیقی (میلیارد تومان)",
                  "میانگین قدرت خریدار",
              ],
          ),
          use_container_width=True,
      )

  with tab3:
    st.subheader("🔍 دیدبان پیشرفته و پول هوشمند")
    min_bp = st.sidebar.slider(
        "حداقل قدرت خریدار (Buyer Power):", 0.5, 5.0, 1.0, 0.1
    )
    if "Buyer_Power" in df_filtered.columns:
      df_filtered = df_filtered[df_filtered["Buyer_Power"] >= min_bp]

    s_query = st.text_input("🔍 جستجوی نماد یا نام:", "")
    if s_query:
      q_norm = normalize_persian(s_query)
      m_search = pd.Series(False, index=df_filtered.index)
      for col in df_filtered.columns:
        if df_filtered[col].dtype == "object" or col == "Ticker":
          m_search = m_search | df_filtered[col].astype(str).apply(
              normalize_persian
          ).str.contains(q_norm, na=False)
      df_filtered = df_filtered[m_search]

    col_map = {
        "Ticker": "نماد",
        "Name": "نام شرکت",
        "Sector": "صنعت",
        "Close": "آخرین قیمت",
        "Close(%)": "درصد آخرین",
        "Volume": "حجم معاملات",
        "Value_Bln": "ارزش معاملات (ب.ت)",
        "Buyer_Power": "قدرت خریدار",
        "Net_Real_Flow": "خالص پول حقیقی (ب.ت)",
    }
    avail_cols = [
        col for col in col_map.keys() if col in df_filtered.columns
    ]
    df_view = df_filtered[avail_cols].rename(columns=col_map)
    st.dataframe(
        df_view.style.applymap(
            color_surplus_df,
            subset=["درصد آخرین", "قدرت خریدار", "خالص پول حقیقی (ب.ت)"],
        ),
        use_container_width=True,
    )

  with tab4:
    st.subheader("📊 برترین صف‌های خرید و فروش بازار")
    sc1, sc2 = st.columns(2)
    with sc1:
      st.markdown("### 🔥 برترین صف‌های خرید")
      df_bq_t = df_filtered[df_filtered["Clean_BQ_Value"] > 0].sort_values(
          by="Clean_BQ_Value", ascending=False
      )
      if not df_bq_t.empty:
        st.dataframe(
            df_bq_t[["Ticker", "Clean_BQ_Value", "Close(%)"]]
            .rename(
                columns={
                    "Ticker": "نماد",
                    "Clean_BQ_Value": "ارزش صف (ب.ت)",
                    "Close(%)": "درصد آخرین",
                }
            )
            .style.applymap(color_surplus_df, subset=["درصد آخرین"]),
            use_container_width=True,
        )
      else:
        st.info("صف خریدی ثبت نشده است.")

    with sc2:
      st.markdown("### ❄️ برترین صف‌های فروش")
      df_sq_t = df_filtered[df_filtered["Clean_SQ_Value"] > 0].sort_values(
          by="Clean_SQ_Value", ascending=False
      )
      if not df_sq_t.empty:
        st.dataframe(
            df_sq_t[["Ticker", "Clean_SQ_Value", "Close(%)"]]
            .rename(
                columns={
                    "Ticker": "نماد",
                    "Clean_SQ_Value": "ارزش صف (ب.ت)",
                    "Close(%)": "درصد آخرین",
                }
            )
            .style.applymap(color_surplus_df, subset=["درصد آخرین"]),
            use_container_width=True,
        )
      else:
        st.info("صف فروشی ثبت نشده است.")

  with tab5:
    st.subheader("🤖 استراتژیست هوشمند بازار")
    sig_col1, sig_col2 = st.columns(2)
    with sig_col1:
      st.markdown("### 🚀 سیگنال خرید طلایی (ورود پول هوشمند)")
      if (
          "Buyer_Power" in df_filtered.columns
          and "Net_Real_Flow" in df_filtered.columns
      ):
        df_buy_signal = df_filtered[
            (df_filtered["Buyer_Power"] >= 1.5)
            & (df_filtered["Net_Real_Flow"] > 0.5)
            & (df_filtered["Close(%)"] > 0)
        ].sort_values(by="Buyer_Power", ascending=False)

        if not df_buy_signal.empty:
          st.dataframe(
              df_buy_signal[
                  [
                      "Ticker",
                      "Sector",
                      "Close(%)",
                      "Buyer_Power",
                      "Net_Real_Flow",
                  ]
              ]
              .rename(
                  columns={
                      "Ticker": "نماد",
                      "Sector": "صنعت",
                      "Close(%)": "درصد تغییر",
                      "Buyer_Power": "قدرت خریدار",
                      "Net_Real_Flow": "ورود پول (ب.ت)",
                  }
              )
              .style.applymap(
                  color_surplus_df,
                  subset=["درصد تغییر", "قدرت خریدار", "ورود پول (ب.ت)"],
              ),
              use_container_width=True,
          )
        else:
          st.info("نمادی با شرایط سخت‌گیرانه خرید طلایی یافت نشد.")

    with sig_col2:
      st.markdown("### ⚠️ هشدار و سیگنال خروج (خروج پول حقیقی)")
      if (
          "Buyer_Power" in df_filtered.columns
          and "Net_Real_Flow" in df_filtered.columns
      ):
        df_sell_signal = df_filtered[
            (df_filtered["Buyer_Power"] <= 0.7)
            & (df_filtered["Net_Real_Flow"] < -0.2)
        ].sort_values(by="Net_Real_Flow", ascending=True)

        if not df_sell_signal.empty:
          st.dataframe(
              df_sell_signal[
                  [
                      "Ticker",
                      "Sector",
                      "Close(%)",
                      "Buyer_Power",
                      "Net_Real_Flow",
                  ]
              ]
              .rename(
                  columns={
                      "Ticker": "نماد",
                      "Sector": "صنعت",
                      "Close(%)": "درصد تغییر",
                      "Buyer_Power": "قدرت خریدار",
                      "Net_Real_Flow": "خروج پول (ب.ت)",
                  }
              )
              .style.applymap(
                  color_surplus_df,
                  subset=["درصد تغییر", "قدرت خریدار", "خروج پول (ب.ت)"],
              ),
              use_container_width=True,
          )
        else:
          st.info("هشدار جدی برای خروج پول سنگین گزارش نشده است.")

  with tab6:
    st.subheader("💼 تحلیل هوشمند پورتفوی شخصی")
    st.markdown(
        "در این بخش می‌توانید نمادهای پورتفوی خود را به‌صورت خودکار (از طریق فایل"
        " اکسل آپلود شده) یا دستی تحلیل کنید تا وضعیت نگهداری، خرید مجدد یا"
        " فروش مشخص شود."
    )

    user_symbols = []
    if portfolio_file is not None:
      try:
        df_port = pd.read_excel(portfolio_file)
        col_ticker_port = next(
            (c for c in ["نام نماد", "Ticker", "نماد"] if c in df_port.columns),
            None,
        )
        if col_ticker_port:
          user_symbols = (
              df_port[col_ticker_port].dropna().astype(str).str.strip().tolist()
          )
          st.success(
              f"تعداد {len(user_symbols)} نماد از فایل پورتفو با موفقیت بارگذاری"
              " شد."
          )
      except Exception as e:
        st.error(f"در خواندن فایل پورتفو خطایی رخ داد: {e}")

    manual_input = st.text_input(
        "یا نمادهای پورتفو را با کاما (,) وارد کنید (مثلا: خساپا, وبملت, ذوب):",
        "",
    )
    if manual_input:
      manual_list = [s.strip() for s in manual_input.split(",") if s.strip()]
      user_symbols.extend(manual_list)

    user_symbols = list(set(user_symbols))

    if user_symbols:
      df_port_analysis = df_filtered[
          df_filtered["Ticker"].isin(user_symbols)
          | df_filtered["Name"].isin(user_symbols)
      ].copy()

      if not df_port_analysis.empty:

        def get_portfolio_signal(row):
          bp = row.get("Buyer_Power", 1.0)
          flow = row.get("Net_Real_Flow", 0.0)
          bq = row.get("Clean_BQ_Value", 0.0)
          sq = row.get("Clean_SQ_Value", 0.0)
          chg = row.get("Close(%)", 0.0)

          if sq > 0 or (bp < 0.7 and flow < 0):
            return "🔴 سیگنال فروش / سیو سود (فشار عرضه)"
          elif bq > 0 or (bp >= 1.5 and flow > 0):
            return "🟢 خرید مجدد / اضافه کردن (ورود پول)"
          elif chg > 0:
            return "🟡 نگهداری (روند مثبت آرام)"
          else:
            return "🟡 نگهداری / بررسی پویایی سهم"

        df_port_analysis["سیگنال پورتفو"] = df_port_analysis.apply(
            get_portfolio_signal, axis=1
        )

        view_port_cols = [
            "Ticker",
            "Sector",
            "Close(%)",
            "Buyer_Power",
            "Net_Real_Flow",
            "سیگنال پورتفو",
        ]
        avail_port_cols = [
            c for c in view_port_cols if c in df_port_analysis.columns
        ]

        st.markdown("### 📊 وضعیت لحظه‌ای و سیگنال‌دهی نمادهای پورتفوی شما")
        st.dataframe(
            df_port_analysis[avail_port_cols]
            .rename(
                columns={
                    "Ticker": "نماد",
                    "Sector": "صنعت",
                    "Close(%)": "درصد آخرین",
                    "Buyer_Power": "قدرت خریدار",
                    "Net_Real_Flow": "خالص پول (ب.ت)",
                }
            )
            .style.applymap(
                color_surplus_df,
                subset=["درصد آخرین", "قدرت خریدار", "خالص پول (ب.ت)"],
            ),
            use_container_width=True,
        )
      else:
        st.warning(
            "اطلاعاتی برای نمادهای وارد شده در تابلوی امروز بازار یافت نشد."
        )
    else:
      st.info(
          "لطفاً فایل اکسل پورتفوی خود را آپلود کنید یا نمادهای خود را به صورت"
          " دستی در کادر بالا وارد نمایید."
      )

else:
  st.warning("⚠️ داده‌ای دریافت نشد یا ارتباط با سرور برقرار نیست.")