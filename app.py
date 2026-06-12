import os
import glob
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# ------------------------------------------------------------------
# Resolve paths relative to THIS script (Streamlit Cloud safe)
# ------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------------
# Page config
# ------------------------------------------------------------------
st.set_page_config(page_title="InCites 대학 연구성과 대시보드", layout="wide")

# ------------------------------------------------------------------
# Constants
# ------------------------------------------------------------------
HIGHLIGHT_NAME = "명지대학교"
COLOR_HIGHLIGHT = "#E8392A"
COLOR_OTHER     = "#93B8E0"
COLOR_BG        = "#F5F7FA"
COLOR_SIDEBAR   = "#0D2B5E"
COLOR_MEAN      = "#555555"

DATA_FILE = "Incites_data_example.xlsx"

# Column keys
C_NAME   = "대학명"
C_DOCS   = "논문수"
C_CITED  = "피인용논문비율"
C_HCITE  = "고피인용논문비율"
C_CNCI   = "CNCI"
C_TOP10  = "상위10%논문비율"

ALL_METRICS = [C_DOCS, C_CITED, C_HCITE, C_CNCI, C_TOP10]

METRIC_LABEL = {
    C_DOCS:  "논문수 (WoS Documents)",
    C_CITED: "피인용논문비율 (% Docs Cited)",
    C_HCITE: "고피인용논문비율 (% Highly Cited)",
    C_CNCI:  "CNCI (범주정규화 피인용지수)",
    C_TOP10: "상위 10% 논문비율 (% Top 10%)",
}

METRIC_DESC = {
    C_DOCS:  "Web of Science에 등재된 총 논문 수. 연구 생산성 규모를 나타냅니다.",
    C_CITED: "전체 논문 중 1회 이상 인용된 논문의 비율. 연구 활용·확산 정도를 나타냅니다.",
    C_HCITE: "전 세계 동일 분야 상위 1% 피인용 논문에 해당하는 비율. 최상위 연구 영향력을 나타냅니다.",
    C_CNCI:  "Category Normalized Citation Impact. 분야·연도·문헌유형 정규화 피인용 지수 (1.0 = 세계 평균).",
    C_TOP10: "전 세계 동일 분야 상위 10% 피인용 논문에 해당하는 비율. 연구의 질적 우수성을 나타냅니다.",
}

METRIC_UNIT = {
    C_DOCS:  "정수 (편)",
    C_CITED: "비율 (0~1)",
    C_HCITE: "비율 (0~1)",
    C_CNCI:  "1.0 = 세계 평균",
    C_TOP10: "비율 (0~1)",
}

FMT_PLOTLY = {C_DOCS: ",.0f", C_CITED: ".2%", C_HCITE: ".2%", C_CNCI: ".2f", C_TOP10: ".2%"}

def fmt_val(col, v):
    if col == C_DOCS:  return f"{v:,.0f}"
    if col == C_CNCI:  return f"{v:.2f}"
    return f"{v:.2%}"

def fmt_style(col):
    if col == C_DOCS:  return "{:,.0f}"
    if col == C_CNCI:  return "{:.2f}"
    return "{:.2%}"

def color_list(names):
    return [COLOR_HIGHLIGHT if n == HIGHLIGHT_NAME else COLOR_OTHER for n in names]

# ------------------------------------------------------------------
# Global CSS
# ------------------------------------------------------------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
html, body, [class*="css"], .stApp {{
    font-family: 'Noto Sans KR', sans-serif !important;
}}
.stApp {{ background-color: {COLOR_BG}; }}
section[data-testid="stSidebar"] {{ background-color: {COLOR_SIDEBAR}; }}
section[data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
.footnote {{
    color: #8A94A6; font-size: 0.78rem;
    border-top: 1px solid #D7DEE8;
    padding-top: 8px; margin-top: 32px; line-height: 1.6;
}}
h1, h2, h3 {{ color: {COLOR_SIDEBAR}; }}
div[data-testid="stMetric"] {{
    background: #FFFFFF; border-radius: 10px;
    padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# Data loading (robust: auto-discover xlsx)
# ------------------------------------------------------------------
@st.cache_data
def load_data():
    path = os.path.join(BASE_DIR, DATA_FILE)
    if not os.path.exists(path):
        candidates = sorted(glob.glob(os.path.join(BASE_DIR, "*.xlsx")))
        if not candidates:
            st.error(
                f"엑셀 파일을 찾을 수 없습니다. '{DATA_FILE}'이 app.py와 같은 "
                f"폴더에 있는지 확인하세요.\n\n현재 폴더: {os.listdir(BASE_DIR)}"
            )
            st.stop()
        path = candidates[0]

    df = pd.read_excel(path, sheet_name=0)
    df[C_NAME] = df[C_NAME].astype(str).str.strip()
    for c in ALL_METRICS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.dropna(subset=[C_NAME]).reset_index(drop=True)

df = load_data()

# ------------------------------------------------------------------
# Sidebar
# ------------------------------------------------------------------
st.sidebar.markdown("## 📊 InCites 대시보드")
st.sidebar.markdown("Web of Science · 한국 대학 연구성과")
st.sidebar.markdown("---")
page = st.sidebar.radio(
    "페이지 선택",
    ["📋 데이터 개요", "🏆 전체 순위", "🔬 심층 분석", "🎯 우리 대학 프로파일"],
)
st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<div style='background:rgba(255,255,255,0.1);border-radius:8px;"
    f"padding:10px;text-align:center;'>"
    f"<span style='font-size:0.75rem;'>강조 대학</span><br>"
    f"<b style='color:{COLOR_HIGHLIGHT};font-size:1.05rem;'>{HIGHLIGHT_NAME}</b></div>",
    unsafe_allow_html=True,
)

def footnote(text):
    st.markdown(f"<div class='footnote'>{text}</div>", unsafe_allow_html=True)

# Helper: standard horizontal bar
def hbar(metric, height=860):
    d = df.sort_values(metric, ascending=True)
    mean_v = df[metric].mean()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=d[C_NAME], x=d[metric], orientation="h",
        marker_color=color_list(d[C_NAME]),
        hovertemplate="%{y}<br>%{x:" + FMT_PLOTLY[metric] + "}<extra></extra>",
    ))
    fig.add_vline(x=mean_v, line_dash="dash", line_color=COLOR_MEAN,
                  annotation_text=f"평균 {fmt_val(metric, mean_v)}",
                  annotation_position="top")
    fig.update_layout(
        height=height, margin=dict(l=10, r=20, t=40, b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title=METRIC_LABEL[metric], yaxis_title="",
        font=dict(family="Noto Sans KR"),
    )
    if metric != C_DOCS and metric != C_CNCI:
        fig.update_xaxes(tickformat=".0%")
    return fig


# ==================================================================
# PAGE 1 — 데이터 개요
# ==================================================================
if page == "📋 데이터 개요":
    st.title("데이터 개요")
    st.markdown("Web of Science **InCites** 데이터셋 기반 한국 대학 연구성과 현황")

    # KPI row
    cols = st.columns(5)
    kpi = [
        ("대학 수",       f"{len(df)} 개교"),
        ("평균 논문수",    f"{df[C_DOCS].mean():,.0f} 편"),
        ("평균 피인용비율", f"{df[C_CITED].mean():.2%}"),
        ("평균 CNCI",     f"{df[C_CNCI].mean():.2f}"),
        ("평균 상위10%",   f"{df[C_TOP10].mean():.2%}"),
    ]
    for c, (label, val) in zip(cols, kpi):
        c.metric(label, val)

    # Metric descriptions
    st.subheader("지표 설명")
    desc_df = pd.DataFrame({
        "지표": [METRIC_LABEL[m] for m in ALL_METRICS],
        "설명": [METRIC_DESC[m] for m in ALL_METRICS],
        "단위 / 해석": [METRIC_UNIT[m] for m in ALL_METRICS],
    })
    st.table(desc_df)

    # Raw data
    st.subheader("원본 데이터 미리보기")
    style_fmt = {m: fmt_style(m) for m in ALL_METRICS}
    st.dataframe(
        df.style.format(style_fmt),
        use_container_width=True, height=460,
    )

    footnote(
        "※ 출처: Web of Science InCites Dataset. "
        "CNCI 1.0은 전 세계 평균이며, 1.0 이상이면 세계 평균을 상회합니다. "
        "피인용논문비율은 1회 이상 인용된 논문 비중, "
        "고피인용논문비율은 상위 1% 피인용 논문 비중, "
        "상위 10% 논문비율은 상위 10% 피인용 논문 비중입니다."
    )


# ==================================================================
# PAGE 2 — 전체 순위
# ==================================================================
elif page == "🏆 전체 순위":
    st.title("전체 순위")
    st.markdown("지표별 전체 대학 수평 막대 차트입니다. **탭**으로 지표를 전환하세요.")

    tab_labels = [METRIC_LABEL[m] for m in ALL_METRICS]
    tabs = st.tabs(tab_labels)

    for tab, metric in zip(tabs, ALL_METRICS):
        with tab:
            st.plotly_chart(hbar(metric), use_container_width=True)

    footnote(
        f"※ 빨간색 막대는 강조 대학({HIGHLIGHT_NAME}), 파란색은 나머지 대학입니다. "
        "점선은 전체 평균을 나타냅니다. 출처: Web of Science InCites."
    )


# ==================================================================
# PAGE 3 — 심층 분석
# ==================================================================
elif page == "🔬 심층 분석":
    st.title("심층 분석")
    st.markdown(
        "논문 규모(논문수)와 피인용논문비율의 관계, 그리고 피인용논문비율 분포를 살펴봅니다."
    )

    col1, col2 = st.columns(2)

    # ---- Scatter: 논문수 vs 피인용논문비율 ----
    with col1:
        st.subheader("논문수 vs 피인용논문비율 산점도")
        d = df.copy()
        d["구분"] = np.where(d[C_NAME] == HIGHLIGHT_NAME, HIGHLIGHT_NAME, "기타 대학")

        fig = px.scatter(
            d, x=C_DOCS, y=C_CITED, color="구분",
            color_discrete_map={HIGHLIGHT_NAME: COLOR_HIGHLIGHT, "기타 대학": COLOR_OTHER},
            hover_name=C_NAME,
            hover_data={C_DOCS: ":,.0f", C_CITED: ":.2%", "구분": False},
        )
        fig.update_traces(marker=dict(size=12, line=dict(width=0.5, color="white")))

        # Mean lines
        fig.add_hline(y=df[C_CITED].mean(), line_dash="dash", line_color=COLOR_MEAN,
                      annotation_text=f"피인용 평균 {df[C_CITED].mean():.2%}",
                      annotation_position="top left")
        fig.add_vline(x=df[C_DOCS].mean(), line_dash="dash", line_color=COLOR_MEAN,
                      annotation_text=f"논문수 평균 {df[C_DOCS].mean():,.0f}",
                      annotation_position="top right")

        # Quadrant labels
        x_mid, y_mid = df[C_DOCS].mean(), df[C_CITED].mean()
        x_max, y_max = df[C_DOCS].max(), df[C_CITED].max()
        x_min, y_min = df[C_DOCS].min(), df[C_CITED].min()
        quad = [
            (x_max * 0.80, y_max * 0.99, "규모↑ 영향력↑"),
            (x_min + (x_mid - x_min) * 0.15, y_max * 0.99, "규모↓ 영향력↑"),
            (x_max * 0.80, y_min + (y_mid - y_min) * 0.15, "규모↑ 영향력↓"),
            (x_min + (x_mid - x_min) * 0.15, y_min + (y_mid - y_min) * 0.15, "규모↓ 영향력↓"),
        ]
        for qx, qy, qtxt in quad:
            fig.add_annotation(x=qx, y=qy, text=qtxt, showarrow=False,
                               font=dict(size=10, color="#AAAAAA"))

        fig.update_layout(
            height=540, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Noto Sans KR"), legend_title_text="",
            xaxis_title="논문수", yaxis_title="피인용논문비율",
        )
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    # ---- Histogram: 피인용논문비율 분포 ----
    with col2:
        st.subheader("피인용논문비율 분포 히스토그램")
        fig2 = go.Figure()
        fig2.add_trace(go.Histogram(
            x=df[C_CITED], nbinsx=15,
            marker_color=COLOR_OTHER, opacity=0.85, name="대학 분포",
        ))
        # Highlight line
        hv = df.loc[df[C_NAME] == HIGHLIGHT_NAME, C_CITED]
        if not hv.empty:
            fig2.add_vline(
                x=hv.iloc[0], line_color=COLOR_HIGHLIGHT, line_width=3,
                annotation_text=f"{HIGHLIGHT_NAME} {hv.iloc[0]:.2%}",
                annotation_position="top right",
            )
        # Mean line
        fig2.add_vline(
            x=df[C_CITED].mean(), line_dash="dash", line_color=COLOR_MEAN,
            annotation_text=f"평균 {df[C_CITED].mean():.2%}",
            annotation_position="bottom left",
        )
        fig2.update_layout(
            height=540, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Noto Sans KR"), bargap=0.05,
            xaxis_title="피인용논문비율", yaxis_title="대학 수",
        )
        fig2.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig2, use_container_width=True)

    footnote(
        "※ 산점도의 점선은 각 지표의 전체 평균을 기준으로 사분면을 형성합니다. "
        "우상단은 '규모와 영향력을 모두 갖춘' 영역입니다. "
        "히스토그램의 빨간 실선은 강조 대학의 위치입니다. "
        "출처: Web of Science InCites."
    )


# ==================================================================
# PAGE 4 — 우리 대학 프로파일
# ==================================================================
else:
    st.title("우리 대학 프로파일")

    default_idx = (
        int(df.index[df[C_NAME] == HIGHLIGHT_NAME][0])
        if (df[C_NAME] == HIGHLIGHT_NAME).any() else 0
    )
    target = st.selectbox("대학 선택", df[C_NAME].tolist(), index=default_idx)
    row = df[df[C_NAME] == target].iloc[0]
    total = len(df)
    is_highlight = (target == HIGHLIGHT_NAME)

    # ---- 5-metric rank summary ----
    st.subheader(f"📌 {target} — 지표별 순위 요약")
    cols = st.columns(5)
    for col_metric, container in zip(ALL_METRICS, cols):
        rank = int(df[col_metric].rank(ascending=False, method="min")
                   [df[C_NAME] == target].iloc[0])
        container.metric(
            METRIC_LABEL[col_metric],
            fmt_val(col_metric, row[col_metric]),
            f"{rank}위 / {total}개교",
        )

    st.markdown("")
    col_a, col_b = st.columns([1, 1.3])

    # ---- Radar chart (percentile) ----
    with col_a:
        st.subheader("레이더 차트 (백분위)")

        cats = [METRIC_LABEL[m] for m in ALL_METRICS]
        pct = [
            df[m].rank(pct=True)[df[C_NAME] == target].iloc[0] * 100
            for m in ALL_METRICS
        ]
        # Close the polygon
        cats_closed = cats + [cats[0]]
        pct_closed  = pct  + [pct[0]]

        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=pct_closed, theta=cats_closed, fill="toself",
            fillcolor=(COLOR_HIGHLIGHT if is_highlight else COLOR_OTHER) + "33",
            line_color=COLOR_HIGHLIGHT if is_highlight else COLOR_OTHER,
            line_width=2.5,
            name=target,
        ))
        # 50th percentile reference
        fig.add_trace(go.Scatterpolar(
            r=[50]*6, theta=cats_closed,
            line=dict(color=COLOR_MEAN, width=1, dash="dot"),
            showlegend=False, hoverinfo="skip",
        ))
        fig.update_layout(
            height=480, paper_bgcolor="rgba(0,0,0,0)",
            polar=dict(
                radialaxis=dict(range=[0, 100], ticksuffix="%", tickfont=dict(size=10)),
                angularaxis=dict(tickfont=dict(size=11)),
            ),
            font=dict(family="Noto Sans KR"), showlegend=False,
            margin=dict(l=60, r=60, t=40, b=30),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.caption(
            "각 지표를 전체 대학 대비 백분위로 환산 (100%에 가까울수록 우수). "
            "점선 원은 50번째 백분위(중앙값)입니다."
        )

    # ---- Top-20 comparison ----
    with col_b:
        st.subheader("상위 20개교 비교")
        metric_sel = st.selectbox(
            "비교 지표", ALL_METRICS,
            format_func=lambda c: METRIC_LABEL[c],
        )
        top20 = df.sort_values(metric_sel, ascending=False).head(20)
        # Ensure target is visible
        if target not in top20[C_NAME].values:
            top20 = pd.concat([top20, df[df[C_NAME] == target]])
        top20 = top20.sort_values(metric_sel, ascending=True)

        mean_v = df[metric_sel].mean()
        fig = go.Figure(go.Bar(
            y=top20[C_NAME], x=top20[metric_sel], orientation="h",
            marker_color=color_list(top20[C_NAME]),
            hovertemplate="%{y}<br>%{x:" + FMT_PLOTLY[metric_sel] + "}<extra></extra>",
        ))
        fig.add_vline(
            x=mean_v, line_dash="dash", line_color=COLOR_MEAN,
            annotation_text=f"전체 평균 {fmt_val(metric_sel, mean_v)}",
        )
        fig.update_layout(
            height=600, margin=dict(l=10, r=20, t=30, b=10),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Noto Sans KR"),
            xaxis_title=METRIC_LABEL[metric_sel], yaxis_title="",
        )
        if metric_sel not in (C_DOCS, C_CNCI):
            fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    footnote(
        f"※ 순위는 전체 {total}개교 대비 산정되었습니다. "
        "레이더 차트는 절대값이 아닌 백분위 기준이며, 5개 지표를 한눈에 비교할 수 있습니다. "
        "선택 대학이 상위 20위 밖이면 비교 차트에 추가 표시됩니다. "
        "출처: Web of Science InCites."
    )
