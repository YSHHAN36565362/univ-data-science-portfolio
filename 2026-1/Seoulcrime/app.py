import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import folium
from folium.plugins import HeatMap
from streamlit_folium import st_folium
import requests
import json
import numpy as np
import xgboost as xgb
from sklearn.preprocessing import MinMaxScaler
from google import genai

# ============================================================
# 페이지 설정 & 커스텀 CSS (서울안전누리 스타일 다크 테마)
# ============================================================
st.set_page_config(
    page_title="서울 SafeCity – 범죄 위험도 분석 시스템",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 다크 테마 CSS
st.markdown("""
<style>
/* ===== 전체 배경 & 텍스트 ===== */
.stApp {
    background-color: #0a0e1a;
}
section[data-testid="stSidebar"] {
    background-color: #111827;
    border-right: 1px solid #1e293b;
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown li,
section[data-testid="stSidebar"] label {
    color: #94a3b8 !important;
}

/* ===== 헤더 배너 ===== */
.header-banner {
    background: linear-gradient(135deg, #111827 0%, #1a2236 100%);
    border: 1px solid #1e293b;
    border-radius: 12px;
    padding: 20px 28px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.header-left {
    display: flex;
    align-items: center;
    gap: 14px;
}
.header-icon {
    width: 44px;
    height: 44px;
    background: linear-gradient(135deg, #3b82f6, #22d3ee);
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 22px;
    box-shadow: 0 0 20px rgba(59,130,246,0.3);
}
.header-title {
    font-size: 22px;
    font-weight: 800;
    color: #f1f5f9;
    letter-spacing: -0.5px;
}
.header-title span {
    color: #22d3ee;
}
.header-sub {
    font-size: 12px;
    color: #64748b;
    margin-top: 2px;
}
.header-right {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 13px;
    color: #64748b;
}
.live-dot {
    width: 8px;
    height: 8px;
    background: #10b981;
    border-radius: 50%;
    display: inline-block;
    animation: pulse 2s infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 1; box-shadow: 0 0 0 0 rgba(16,185,129,0.4); }
    50% { opacity: 0.7; box-shadow: 0 0 0 6px rgba(16,185,129,0); }
}

/* ===== KPI 카드 ===== */
.kpi-container {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 14px;
    margin-bottom: 20px;
}
.kpi-card {
    background: #1a2236;
    border: 1px solid #1e293b;
    border-radius: 12px;
    padding: 18px 20px;
    position: relative;
    overflow: hidden;
    transition: transform 0.2s, box-shadow 0.2s;
}
.kpi-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 0 20px rgba(34,211,238,0.15);
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 12px 12px 0 0;
}
.kpi-card.red::before { background: #ef4444; }
.kpi-card.orange::before { background: #f97316; }
.kpi-card.green::before { background: #10b981; }
.kpi-card.blue::before { background: #3b82f6; }
.kpi-card.cyan::before { background: #22d3ee; }
.kpi-card.purple::before { background: #a855f7; }
.kpi-label {
    font-size: 12px;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
}
.kpi-value {
    font-size: 26px;
    font-weight: 800;
    letter-spacing: -1px;
}
.kpi-value.red { color: #ef4444; }
.kpi-value.orange { color: #f97316; }
.kpi-value.green { color: #10b981; }
.kpi-value.blue { color: #3b82f6; }
.kpi-value.cyan { color: #22d3ee; }
.kpi-value.purple { color: #a855f7; }
.kpi-sub {
    font-size: 11px;
    color: #475569;
    margin-top: 2px;
}

/* ===== 패널 카드 ===== */
.panel-card {
    background: #1a2236;
    border: 1px solid #1e293b;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 16px;
}
.panel-title {
    font-size: 15px;
    font-weight: 700;
    color: #f1f5f9;
    margin-bottom: 14px;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* ===== 랭킹 아이템 ===== */
.rank-item {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 12px;
    border-radius: 8px;
    margin-bottom: 4px;
    background: rgba(255,255,255,0.02);
    transition: background 0.2s;
}
.rank-item:hover {
    background: rgba(59,130,246,0.08);
}
.rank-num {
    width: 26px;
    height: 26px;
    border-radius: 7px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 12px;
    font-weight: 700;
    flex-shrink: 0;
}
.rank-num.danger { background: rgba(239,68,68,0.15); color: #ef4444; }
.rank-num.safe { background: rgba(16,185,129,0.15); color: #10b981; }
.rank-name {
    font-size: 13px;
    font-weight: 600;
    color: #e2e8f0;
    min-width: 70px;
}
.rank-bar {
    flex: 1;
    height: 6px;
    background: rgba(255,255,255,0.05);
    border-radius: 3px;
    overflow: hidden;
}
.rank-bar-fill {
    height: 100%;
    border-radius: 3px;
    transition: width 0.5s;
}
.rank-value {
    font-size: 13px;
    font-weight: 700;
    color: #94a3b8;
    min-width: 50px;
    text-align: right;
}

/* ===== 예측 시뮬레이터 결과 ===== */
.sim-result-card {
    background: rgba(34,211,238,0.05);
    border: 1px solid rgba(34,211,238,0.2);
    border-radius: 12px;
    padding: 24px;
    text-align: center;
}
.sim-result-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
    margin-top: 16px;
}
.sim-result-item {
    background: rgba(255,255,255,0.03);
    border-radius: 8px;
    padding: 16px;
}
.sim-result-label {
    font-size: 11px;
    color: #64748b;
    margin-bottom: 4px;
}
.sim-result-value {
    font-size: 22px;
    font-weight: 800;
}
.sim-note {
    font-size: 12px;
    color: #64748b;
    margin-top: 16px;
    line-height: 1.7;
    text-align: left;
}

/* ===== 필요도 아이템 ===== */
.need-item {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 12px 14px;
    border-radius: 8px;
    margin-bottom: 6px;
    background: rgba(255,255,255,0.02);
    border-left: 3px solid transparent;
}
.need-item.high { border-left-color: #ef4444; }
.need-item.mid { border-left-color: #f97316; }
.need-item.low { border-left-color: #eab308; }

/* ===== Plotly 차트 배경 투명 ===== */
.stPlotlyChart {
    background: transparent !important;
}

/* ===== 기본 st.metric 숨기고 커스텀 사용 ===== */
[data-testid="stMetric"] {
    background: #1a2236;
    border: 1px solid #1e293b;
    border-radius: 10px;
    padding: 12px 16px;
}

/* ===== 사이드바 제목 ===== */
.sidebar-logo {
    text-align: center;
    padding: 10px 0 20px 0;
    border-bottom: 1px solid #1e293b;
    margin-bottom: 20px;
}
.sidebar-logo-text {
    font-size: 20px;
    font-weight: 800;
    color: #f1f5f9;
}
.sidebar-logo-text span {
    color: #22d3ee;
}

/* ===== 구분선 ===== */
.divider {
    border: none;
    border-top: 1px solid #1e293b;
    margin: 16px 0;
}

/* ===== 탭 스타일 오버라이드 ===== */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #111827;
    border-radius: 8px;
    padding: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 6px;
    color: #64748b;
    font-weight: 600;
}
.stTabs [aria-selected="true"] {
    background: #3b82f6 !important;
    color: white !important;
}

/* ===== 반응형 KPI ===== */
@media (max-width: 768px) {
    .kpi-container {
        grid-template-columns: repeat(2, 1fr);
    }
    .sim-result-grid {
        grid-template-columns: 1fr;
    }
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# 데이터 로딩 (기존 구조 유지)
# ============================================================
@st.cache_data(ttl=0)
def load_data():
    # --- 범죄율/검거율 (2019~2023) ---
    try:
        crime = pd.read_csv('자치구별 범죄율 검거율 5개년.csv', encoding='utf-8', header=1)
    except:
        crime = pd.read_csv('자치구별 범죄율 검거율 5개년.csv', encoding='cp949', header=1)
    crime.columns = ['자치구별','2019_범죄율','2019_검거율','2020_범죄율','2020_검거율',
                     '2021_범죄율','2021_검거율','2022_범죄율','2022_검거율','2023_범죄율','2023_검거율']
    crime = crime.dropna(subset=['자치구별'])

    # --- 발생/검거 건수 (2019~2023) ---
    try:
        occur = pd.read_csv('전국 발생 검거 수.csv', encoding='utf-8', header=1)
    except:
        occur = pd.read_csv('전국 발생 검거 수.csv', encoding='cp949', header=1)
    occur.columns = ['자치구별','2019_발생','2019_검거','2020_발생','2020_검거',
                     '2021_발생','2021_검거','2022_발생','2022_검거','2023_발생','2023_검거']
    occur = occur.dropna(subset=['자치구별'])
    occur = occur[occur['자치구별'] != '소계']
    for col in occur.columns[1:]:
        occur[col] = pd.to_numeric(occur[col].astype(str).str.replace(',',''), errors='coerce')

    # --- 2024년 범죄 데이터 추가 (crime_seoul.csv) ---
    try:
        crime_2024_raw = pd.read_csv('crime_seoul.csv', encoding='utf-8')
        crime_2024_list = []
        for _, row in crime_2024_raw.iloc[4:].iterrows():  # skip header rows + 소계
            gu = str(row.iloc[1]).strip()
            if gu == '소계' or gu == '자치구별(2)':
                continue
            oc = int(str(row.iloc[2]).replace(',',''))
            ar = int(str(row.iloc[3]).replace(',',''))
            crime_2024_list.append({'자치구별': gu, '2024_발생': oc, '2024_검거': ar})
        crime_2024_df = pd.DataFrame(crime_2024_list)
        # 발생/검거 건수에 2024년 추가
        occur = pd.merge(occur, crime_2024_df, on='자치구별', how='left')
    except Exception as e:
        st.warning(f"crime_seoul.csv 로딩 실패: {e}")

    # --- CCTV ---
    cctv = pd.read_csv('cctv_clean.csv', encoding='utf-8-sig')
    cctv.columns = ['자치구','총계','2016년이전','2017년','2018년','2019년','2020년',
                    '2021년','2022년','2023년','2024년','2025년']
    cctv = cctv[cctv['자치구'] != '계']
    cctv = cctv.dropna(subset=['자치구'])
    cctv['자치구'] = cctv['자치구'].str.replace(' ', '')
    for col in cctv.columns[1:]:
        cctv[col] = pd.to_numeric(cctv[col].astype(str).str.replace(',',''), errors='coerce')

    # --- 인구 (2019~2023) ---
    try:
        pop = pd.read_csv('인구 수.csv', encoding='utf-8', header=1)
    except:
        pop = pd.read_csv('인구 수.csv', encoding='cp949', header=1)
    pop.columns = ['자치구별','2019_인구','2020_인구','2021_인구','2022_인구','2023_인구']
    pop = pop.dropna(subset=['자치구별'])
    pop = pop[pop['자치구별'] != '서울특별시']
    for col in pop.columns[1:]:
        pop[col] = pd.to_numeric(pop[col].astype(str).str.replace(',',''), errors='coerce')

    # --- 2025년 인구 데이터 추가 (population_seoul.csv) → 2024년 분석에 사용 ---
    try:
        # 인코딩 자동 감지
        for enc in ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr']:
            try:
                pop_2025_raw = pd.read_csv('population_seoul.csv', encoding=enc)
                break
            except:
                continue
        pop_2025_list = []
        for _, row in pop_2025_raw.iloc[1:].iterrows():
            raw = str(row.iloc[0])
            parts = raw.split()
            gu = parts[1] if len(parts) >= 2 else parts[0]
            population = int(str(row.iloc[1]).replace(',',''))
            pop_2025_list.append({'자치구별': gu, '2024_인구': population})
        pop_2025_df = pd.DataFrame(pop_2025_list)
        pop = pd.merge(pop, pop_2025_df, on='자치구별', how='left')
    except Exception as e:
        st.warning(f"population_seoul.csv 로딩 실패: {e}")

    # --- 2024년 범죄율/검거율 계산하여 crime 테이블에 추가 ---
    try:
        merged_2024 = pd.merge(
            occur[['자치구별','2024_발생','2024_검거']].dropna(),
            pop[['자치구별','2024_인구']].dropna(),
            on='자치구별', how='inner'
        )
        merged_2024['2024_범죄율'] = (merged_2024['2024_발생'] / merged_2024['2024_인구'] * 100).round(2)
        merged_2024['2024_검거율'] = (merged_2024['2024_검거'] / merged_2024['2024_발생'] * 100).round(2)
        crime = pd.merge(crime, merged_2024[['자치구별','2024_범죄율','2024_검거율']], on='자치구별', how='left')
    except Exception as e:
        st.warning(f"2024 범죄율 계산 실패: {e}")

    # --- 면적 데이터 (area_seoul.csv) ---
    area = pd.DataFrame()
    try:
        for enc in ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr']:
            try:
                area_raw = pd.read_csv('area_seoul.csv', encoding=enc)
                break
            except:
                continue
        area_list = []
        for _, row in area_raw.iloc[2:].iterrows():
            gu = str(row.iloc[1]).strip()
            if gu == '소계':
                continue
            km2 = float(str(row.iloc[2]).replace(',',''))
            area_list.append({'자치구별': gu, '면적_km2': km2})
        area = pd.DataFrame(area_list)
    except Exception as e:
        st.warning(f"area_seoul.csv 로딩 실패: {e}")

    return crime, occur, cctv, pop, area

crime, occur, cctv, pop, area = load_data()

@st.cache_data
def load_geojson():
    url = "https://raw.githubusercontent.com/southkorea/seoul-maps/master/kostat/2013/json/seoul_municipalities_geo_simple.json"
    return requests.get(url).json()

seoul_geo = load_geojson()

# 자치구 중심 좌표
gu_coords = {
    '종로구':[37.5735,126.9790],'중구':[37.5641,126.9979],'용산구':[37.5326,126.9910],
    '성동구':[37.5634,127.0371],'광진구':[37.5384,127.0822],'동대문구':[37.5744,127.0400],
    '중랑구':[37.6063,127.0928],'성북구':[37.5894,127.0167],'강북구':[37.6397,127.0255],
    '도봉구':[37.6688,127.0471],'노원구':[37.6542,127.0568],'은평구':[37.6027,126.9291],
    '서대문구':[37.5791,126.9368],'마포구':[37.5663,126.9014],'양천구':[37.5170,126.8664],
    '강서구':[37.5510,126.8495],'구로구':[37.4955,126.8876],'금천구':[37.4519,126.8968],
    '영등포구':[37.5264,126.8963],'동작구':[37.5124,126.9393],'관악구':[37.4784,126.9516],
    '서초구':[37.4837,127.0324],'강남구':[37.5172,127.0473],'송파구':[37.5145,127.1050],
    '강동구':[37.5301,127.1238]
}


# ============================================================
# 헬퍼 함수
# ============================================================
def get_risk_level(rate):
    if rate >= 1.5: return '위험', '#ef4444', '🔴'
    if rate >= 1.0: return '주의', '#f97316', '🟠'
    if rate >= 0.7: return '보통', '#eab308', '🟡'
    return '안전', '#10b981', '🟢'

def plotly_dark_layout(fig, height=450):
    """공통 Plotly 다크 테마 레이아웃"""
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8', size=12),
        height=height,
        margin=dict(l=40, r=20, t=40, b=40),
        xaxis=dict(gridcolor='#1e293b', zerolinecolor='#1e293b'),
        yaxis=dict(gridcolor='#1e293b', zerolinecolor='#1e293b'),
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8')),
    )
    return fig

def make_kpi_html(cards):
    """KPI 카드 그리드 HTML 생성"""
    html = '<div class="kpi-container">'
    for label, value, sub, color in cards:
        html += f'''
        <div class="kpi-card {color}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value {color}">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>'''
    html += '</div>'
    return html

def make_ranking_html(title, items, rank_type='danger'):
    """랭킹 리스트 HTML"""
    max_val = items[0][1] if items else 1
    html = f'<div class="panel-title">{title}</div>'
    for i, (name, val) in enumerate(items):
        pct = val / max_val * 100
        bar_color = '#ef4444' if rank_type == 'danger' else '#10b981'
        html += f'''
        <div class="rank-item">
            <div class="rank-num {rank_type}">{i+1}</div>
            <div class="rank-name">{name}</div>
            <div class="rank-bar"><div class="rank-bar-fill" style="width:{pct:.0f}%;background:{bar_color};"></div></div>
            <div class="rank-value">{val}%</div>
        </div>'''
    return html


# ============================================================
# 사이드바
# ============================================================
with st.sidebar:
    st.markdown('''
    <div class="sidebar-logo">
        <div style="font-size:28px;margin-bottom:4px;">🛡️</div>
        <div class="sidebar-logo-text">서울 <span>SafeCity</span></div>
        <div style="font-size:11px;color:#64748b;margin-top:4px;">범죄 위험도 분석 시스템</div>
    </div>
    ''', unsafe_allow_html=True)

    menu = st.radio(
        "메뉴",
        ["🗺️ 안전 지도", "📊 범죄 현황", "📹 CCTV 현황", "📈 통계 비교", "🔮 CCTV 예측 시뮬레이터", "🔍 범죄 유형별 분석", "📉 범죄 추이 예측", "🧠 AI 심층 분석"],
        label_visibility="collapsed"
    )

    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    year = st.selectbox("📅 분석 연도", [2024, 2023, 2022, 2021, 2020, 2019])

    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown(f'''
    <div style="font-size:11px;color:#475569;line-height:1.6;">
        <span class="live-dot"></span> 데이터 업데이트: 2025.12<br>
        학번: 202384064, 202004244, 202384068<br>
        캡스톤디자인 프로젝트
    </div>
    ''', unsafe_allow_html=True)


# ============================================================
# 공통 변수
# ============================================================
col_o = f'{year}_발생'
col_r = f'{year}_범죄율'
col_a = f'{year}_검거율'
col_p = f'{year}_인구'

# 선택한 연도의 컬럼이 데이터에 없으면 가장 가까운 연도로 대체
if col_r not in crime.columns:
    col_r = '2023_범죄율'
if col_a not in crime.columns:
    col_a = '2023_검거율'
if col_o not in occur.columns:
    col_o = '2023_발생'
if col_p not in pop.columns:
    col_p = '2023_인구'


# ============================================================
# 헤더 배너
# ============================================================
st.markdown(f'''
<div class="header-banner">
    <div class="header-left">
        <div class="header-icon">🛡</div>
        <div>
            <div class="header-title">서울 <span>SafeCity</span></div>
            <div class="header-sub">범죄 발생 데이터와 CCTV 정보를 활용한 지역 위험도 예측 시스템</div>
        </div>
    </div>
    <div class="header-right">
        <span class="live-dot"></span> {year}년 데이터 기준
    </div>
</div>
''', unsafe_allow_html=True)


# ============================================================
# 페이지 1: 안전 지도
# ============================================================
if menu == "🗺️ 안전 지도":

    # KPI 카드
    total_crime = int(occur[col_o].sum()) if col_o in occur.columns else 0
    avg_rate = round(crime[col_r].mean(), 2) if col_r in crime.columns else 0
    avg_arrest = round(crime[col_a].mean(), 1) if col_a in crime.columns else 0
    total_cctv = int(cctv['총계'].sum())

    st.markdown(make_kpi_html([
        ('🚨 총 범죄 발생', f'{total_crime:,}건', '서울시 25개 자치구 합계', 'red'),
        ('📊 평균 범죄율', f'{avg_rate}%', '인구 대비 범죄 발생 비율', 'orange'),
        ('🔍 평균 검거율', f'{avg_arrest}%', '범죄 대비 검거 비율', 'green'),
        ('📹 CCTV 총 설치', f'{total_cctv:,}대', '2025년 12월 기준 누적', 'blue'),
    ]), unsafe_allow_html=True)

    # 지도 레이어 선택
    layer = st.radio(
        "지도 레이어",
        ["범죄율", "범죄 발생 건수", "CCTV 밀도", "검거율"],
        horizontal=True,
        label_visibility="collapsed"
    )

    # 지도 + 사이드 패널
    map_col, side_col = st.columns([2, 1])

    with map_col:
        # 다크 테마 지도 생성
        m = folium.Map(
            location=[37.5665, 126.9780],
            zoom_start=11,
            tiles='CartoDB dark_matter'
        )

        # Choropleth 레이어
        if layer == "범죄율" and col_r in crime.columns:
            folium.Choropleth(
                geo_data=seoul_geo, data=crime,
                columns=['자치구별', col_r],
                key_on='feature.properties.name',
                fill_color='YlOrRd', fill_opacity=0.7, line_opacity=0.3,
                legend_name=f'{year}년 범죄율(%)'
            ).add_to(m)
        elif layer == "범죄 발생 건수" and col_o in occur.columns:
            folium.Choropleth(
                geo_data=seoul_geo, data=occur,
                columns=['자치구별', col_o],
                key_on='feature.properties.name',
                fill_color='OrRd', fill_opacity=0.7, line_opacity=0.3,
                legend_name=f'{year}년 범죄 발생 건수'
            ).add_to(m)
        elif layer == "CCTV 밀도":
            folium.Choropleth(
                geo_data=seoul_geo, data=cctv,
                columns=['자치구', '총계'],
                key_on='feature.properties.name',
                fill_color='BuGn', fill_opacity=0.7, line_opacity=0.3,
                legend_name='CCTV 총 설치 대수'
            ).add_to(m)
        elif layer == "검거율" and col_a in crime.columns:
            folium.Choropleth(
                geo_data=seoul_geo, data=crime,
                columns=['자치구별', col_a],
                key_on='feature.properties.name',
                fill_color='Blues', fill_opacity=0.7, line_opacity=0.3,
                legend_name=f'{year}년 검거율(%)'
            ).add_to(m)

        # 자치구별 팝업 마커
        for gu, coord in gu_coords.items():
            rate_val = crime.loc[crime['자치구별']==gu, col_r].values
            rate = float(rate_val[0]) if len(rate_val) > 0 else 0
            risk_label, risk_color, risk_emoji = get_risk_level(rate)

            rate_str = f"{rate_val[0]}%" if len(rate_val) > 0 else "N/A"
            occ_val = occur.loc[occur['자치구별']==gu, col_o].values
            occ_str = f"{int(occ_val[0]):,}건" if len(occ_val) > 0 else "N/A"
            arr_val = crime.loc[crime['자치구별']==gu, col_a].values
            arr_str = f"{arr_val[0]}%" if len(arr_val) > 0 else "N/A"
            cctv_val = cctv.loc[cctv['자치구']==gu, '총계'].values
            cctv_str = f"{int(cctv_val[0]):,}대" if len(cctv_val) > 0 else "N/A"
            pop_val = pop.loc[pop['자치구별']==gu, col_p].values if col_p in pop.columns else []
            pop_str = f"{int(pop_val[0]):,}명" if len(pop_val) > 0 else "N/A"

            popup_html = f"""
            <div style="font-family:-apple-system,sans-serif;width:220px;padding:4px;">
                <div style="font-size:16px;font-weight:700;margin-bottom:6px;">{risk_emoji} {gu}
                    <span style="font-size:11px;background:{risk_color};color:white;padding:2px 8px;border-radius:4px;margin-left:6px;">{risk_label}</span>
                </div>
                <hr style="margin:6px 0;border-color:#eee;">
                <div style="font-size:13px;line-height:1.8;">
                    🚨 범죄 발생: <b>{occ_str}</b><br>
                    📊 범죄율: <b>{rate_str}</b><br>
                    🔍 검거율: <b>{arr_str}</b><br>
                    📹 CCTV: <b>{cctv_str}</b><br>
                    👥 인구: <b>{pop_str}</b>
                </div>
            </div>
            """
            folium.CircleMarker(
                location=coord,
                radius=7,
                color=risk_color,
                fill=True,
                fill_color=risk_color,
                fill_opacity=0.8,
                popup=folium.Popup(popup_html, max_width=260),
                tooltip=f"{gu} ({risk_label})"
            ).add_to(m)

        st_folium(m, width=None, height=550, returned_objects=[])

    with side_col:
        # 위험/안전 랭킹
        if col_r in crime.columns:
            sorted_crime = crime.sort_values(col_r, ascending=False)
            top5 = [(row['자치구별'], row[col_r]) for _, row in sorted_crime.head(5).iterrows()]
            bot5 = [(row['자치구별'], row[col_r]) for _, row in sorted_crime.tail(5).iterrows()]

            st.markdown(f'<div class="panel-card">{make_ranking_html("⚠️ 위험 지역 (범죄율 상위 5)", top5, "danger")}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="panel-card">{make_ranking_html("✅ 안전 지역 (범죄율 하위 5)", bot5, "safe")}</div>', unsafe_allow_html=True)


# ============================================================
# 페이지 2: 범죄 현황
# ============================================================
elif menu == "📊 범죄 현황":

    # KPI
    total_crime = int(occur[col_o].sum()) if col_o in occur.columns else 0
    max_gu = occur.loc[occur[col_o].idxmax(), '자치구별'] if col_o in occur.columns else '-'
    max_val = int(occur[col_o].max()) if col_o in occur.columns else 0
    min_gu = occur.loc[occur[col_o].idxmin(), '자치구별'] if col_o in occur.columns else '-'

    st.markdown(make_kpi_html([
        ('🚨 총 범죄 발생', f'{total_crime:,}건', f'{year}년 서울시 합계', 'red'),
        ('📈 최다 발생 자치구', max_gu, f'{max_val:,}건', 'orange'),
        ('📉 최소 발생 자치구', min_gu, f'{int(occur[col_o].min()):,}건' if col_o in occur.columns else '', 'green'),
        ('📊 평균 범죄율', f'{crime[col_r].mean():.2f}%' if col_r in crime.columns else '-', '25개 자치구 평균', 'blue'),
    ]), unsafe_allow_html=True)

    # 범죄 발생 건수 차트
    if col_o in occur.columns:
        st.markdown('<div class="panel-card"><div class="panel-title">📊 자치구별 범죄 발생 건수</div></div>', unsafe_allow_html=True)
        df_sorted = occur.sort_values(col_o, ascending=True)
        fig1 = px.bar(
            df_sorted, x=col_o, y='자치구별', orientation='h',
            color=col_o, color_continuous_scale='YlOrRd',
            labels={col_o: '발생 건수', '자치구별': ''}
        )
        plotly_dark_layout(fig1, height=600)
        fig1.update_layout(yaxis=dict(dtick=1), coloraxis_showscale=False)
        st.plotly_chart(fig1, use_container_width=True)

    # 범죄율 차트
    if col_r in crime.columns:
        st.markdown('<div class="panel-card"><div class="panel-title">📊 자치구별 범죄율 (%)</div></div>', unsafe_allow_html=True)
        df_sorted2 = crime.sort_values(col_r, ascending=False)
        fig2 = px.bar(
            df_sorted2, x='자치구별', y=col_r,
            color=col_r, color_continuous_scale='OrRd',
            labels={col_r: '범죄율(%)', '자치구별': ''}
        )
        plotly_dark_layout(fig2, height=450)
        fig2.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)

    # 연도별 추이
    st.markdown('<div class="panel-card"><div class="panel-title">📈 연도별 범죄 발생 추이 (상위 5개 구)</div></div>', unsafe_allow_html=True)
    year_cols = ['2019_발생','2020_발생','2021_발생','2022_발생','2023_발생']
    if '2024_발생' in occur.columns:
        year_cols.append('2024_발생')
    occur_years = occur[['자치구별'] + year_cols].copy()
    latest_col = year_cols[-1]
    top5_names = occur_years.nlargest(5, latest_col)['자치구별'].tolist()
    top5_m = occur_years[occur_years['자치구별'].isin(top5_names)].melt(
        id_vars='자치구별', var_name='연도', value_name='발생건수'
    )
    top5_m['연도'] = top5_m['연도'].str.replace('_발생','')
    fig3 = px.line(top5_m, x='연도', y='발생건수', color='자치구별', markers=True)
    plotly_dark_layout(fig3, height=400)
    st.plotly_chart(fig3, use_container_width=True)

    # 검거율
    if col_a in crime.columns:
        st.markdown('<div class="panel-card"><div class="panel-title">🔍 자치구별 검거율 (%)</div></div>', unsafe_allow_html=True)
        fig4 = px.bar(
            crime.sort_values(col_a, ascending=True), x=col_a, y='자치구별',
            orientation='h', color=col_a, color_continuous_scale='Blues',
            labels={col_a: '검거율(%)', '자치구별': ''}
        )
        plotly_dark_layout(fig4, height=600)
        fig4.update_layout(yaxis=dict(dtick=1), coloraxis_showscale=False)
        st.plotly_chart(fig4, use_container_width=True)


# ============================================================
# 페이지 3: CCTV 현황
# ============================================================
elif menu == "📹 CCTV 현황":

    total_cctv = int(cctv['총계'].sum())
    max_cctv_gu = cctv.loc[cctv['총계'].idxmax(), '자치구']
    min_cctv_gu = cctv.loc[cctv['총계'].idxmin(), '자치구']

    # 인구 천명당 CCTV 계산
    cp = cctv[['자치구','총계']].copy().rename(columns={'자치구':'자치구별'})
    cp['자치구별'] = cp['자치구별'].str.strip()
    pp = pop.copy()
    pp['자치구별'] = pp['자치구별'].str.strip()
    mg = pd.merge(cp, pp[['자치구별','2023_인구']], on='자치구별', how='inner')
    mg['인구천명당_CCTV'] = (mg['총계'] / mg['2023_인구'] * 1000).round(2)
    avg_per1k = mg['인구천명당_CCTV'].mean()

    st.markdown(make_kpi_html([
        ('📹 CCTV 총 설치', f'{total_cctv:,}대', '2025년 12월 기준', 'cyan'),
        ('🏆 최다 설치 자치구', max_cctv_gu, f'{int(cctv["총계"].max()):,}대', 'blue'),
        ('📉 최소 설치 자치구', min_cctv_gu, f'{int(cctv["총계"].min()):,}대', 'orange'),
        ('👥 천명당 CCTV 평균', f'{avg_per1k:.1f}대', '인구 1,000명당', 'purple'),
    ]), unsafe_allow_html=True)

    # CCTV 총 설치 대수
    st.markdown('<div class="panel-card"><div class="panel-title">📹 자치구별 CCTV 총 설치 대수</div></div>', unsafe_allow_html=True)
    fig5 = px.bar(
        cctv.sort_values('총계', ascending=False), x='자치구', y='총계',
        color='총계', color_continuous_scale='teal',
        labels={'총계': 'CCTV 대수', '자치구': ''}
    )
    plotly_dark_layout(fig5, height=450)
    fig5.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
    st.plotly_chart(fig5, use_container_width=True)

    # 연도별 CCTV 추이
    st.markdown('<div class="panel-card"><div class="panel-title">📈 연도별 CCTV 설치 추이 (상위 5개 구)</div></div>', unsafe_allow_html=True)
    y_cols = ['자치구','2017년','2018년','2019년','2020년','2021년','2022년','2023년','2024년','2025년']
    top5c = cctv.nlargest(5, '총계')['자치구'].tolist()
    t5m = cctv[cctv['자치구'].isin(top5c)][y_cols].melt(id_vars='자치구', var_name='연도', value_name='설치대수')
    t5m['연도'] = t5m['연도'].str.replace('년','')
    fig6 = px.line(t5m, x='연도', y='설치대수', color='자치구', markers=True)
    plotly_dark_layout(fig6, height=400)
    st.plotly_chart(fig6, use_container_width=True)

    # 인구 천명당 CCTV
    st.markdown('<div class="panel-card"><div class="panel-title">👥 인구 1,000명당 CCTV 설치 비율</div></div>', unsafe_allow_html=True)
    fig7 = px.bar(
        mg.sort_values('인구천명당_CCTV', ascending=False),
        x='자치구별', y='인구천명당_CCTV',
        color='인구천명당_CCTV', color_continuous_scale='Purples',
        labels={'인구천명당_CCTV': '인구 1000명당 CCTV', '자치구별': ''}
    )
    plotly_dark_layout(fig7, height=450)
    fig7.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
    st.plotly_chart(fig7, use_container_width=True)

    # CCTV 설치 현황 지도
    st.markdown('<div class="panel-card"><div class="panel-title">🗺️ CCTV 설치 현황 지도</div></div>', unsafe_allow_html=True)
    m_cctv = folium.Map(location=[37.5665, 126.9780], zoom_start=11, tiles='CartoDB dark_matter')
    folium.Choropleth(
        geo_data=seoul_geo, data=cctv,
        columns=['자치구', '총계'],
        key_on='feature.properties.name',
        fill_color='YlGnBu', fill_opacity=0.7, line_opacity=0.3,
        legend_name='CCTV 설치 대수'
    ).add_to(m_cctv)

    for gu, coord in gu_coords.items():
        cctv_val = cctv.loc[cctv['자치구']==gu, '총계'].values
        if len(cctv_val) > 0:
            folium.CircleMarker(
                location=coord, radius=max(5, cctv_val[0] / 800),
                color='#22d3ee', fill=True, fill_color='#22d3ee', fill_opacity=0.6,
                tooltip=f"{gu}: {int(cctv_val[0]):,}대"
            ).add_to(m_cctv)

    st_folium(m_cctv, width=None, height=500, returned_objects=[])

    # 면적당 CCTV 밀도 (area_seoul.csv 활용)
    if not area.empty:
        st.markdown('<div class="panel-card"><div class="panel-title">📐 면적(km²)당 CCTV 설치 밀도</div></div>', unsafe_allow_html=True)
        area_cctv = pd.merge(
            cctv[['자치구','총계']].rename(columns={'자치구':'자치구별'}),
            area, on='자치구별', how='inner'
        )
        area_cctv['km2당_CCTV'] = (area_cctv['총계'] / area_cctv['면적_km2']).round(1)
        area_cctv = area_cctv.sort_values('km2당_CCTV', ascending=False)

        fig_area_cctv = px.bar(
            area_cctv, x='자치구별', y='km2당_CCTV',
            color='km2당_CCTV', color_continuous_scale='Oranges',
            labels={'km2당_CCTV': 'km²당 CCTV 대수', '자치구별': ''}
        )
        plotly_dark_layout(fig_area_cctv, height=450)
        fig_area_cctv.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
        st.plotly_chart(fig_area_cctv, use_container_width=True)

        # 면적당 범죄 밀도
        if col_o in occur.columns:
            st.markdown('<div class="panel-card"><div class="panel-title">🚨 면적(km²)당 범죄 발생 밀도</div></div>', unsafe_allow_html=True)
            area_crime = pd.merge(
                occur[['자치구별', col_o]],
                area, on='자치구별', how='inner'
            )
            area_crime['km2당_범죄'] = (area_crime[col_o] / area_crime['면적_km2']).round(1)
            area_crime = area_crime.sort_values('km2당_범죄', ascending=False)

            fig_area_crime = px.bar(
                area_crime, x='자치구별', y='km2당_범죄',
                color='km2당_범죄', color_continuous_scale='YlOrRd',
                labels={'km2당_범죄': 'km²당 범죄 건수', '자치구별': ''}
            )
            plotly_dark_layout(fig_area_crime, height=450)
            fig_area_crime.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
            st.plotly_chart(fig_area_crime, use_container_width=True)


# ============================================================
# 페이지 4: 통계 비교
# ============================================================
elif menu == "📈 통계 비교":

    st.markdown('<div class="panel-card"><div class="panel-title">📈 자치구 통계 비교 분석</div></div>', unsafe_allow_html=True)

    gu_list = crime['자치구별'].tolist()
    selected = st.multiselect("비교할 자치구 선택 (최대 5개)", gu_list, default=gu_list[:3], max_selections=5)

    if selected:
        # 비교 테이블
        compare = crime[crime['자치구별'].isin(selected)][['자치구별', col_r, col_a]].copy()
        occ_sel = occur[occur['자치구별'].isin(selected)][['자치구별', col_o]]
        compare = pd.merge(compare, occ_sel, on='자치구별', how='left')
        cctv_sel = cctv[cctv['자치구'].isin(selected)][['자치구','총계']].rename(columns={'자치구':'자치구별'})
        compare = pd.merge(compare, cctv_sel, on='자치구별', how='left')
        pop_sel = pop[pop['자치구별'].isin(selected)][['자치구별', col_p]] if col_p in pop.columns else pd.DataFrame()
        if not pop_sel.empty:
            compare = pd.merge(compare, pop_sel, on='자치구별', how='left')
        compare.columns = ['자치구','범죄율(%)','검거율(%)','발생건수','CCTV대수'] + (['인구'] if not pop_sel.empty else [])

        st.dataframe(compare, use_container_width=True, hide_index=True)

        # 비교 차트
        col1, col2 = st.columns(2)
        with col1:
            fig_r = px.bar(compare, x='자치구', y='범죄율(%)', color='자치구', title='범죄율 비교')
            plotly_dark_layout(fig_r, height=380)
            fig_r.update_layout(showlegend=False)
            st.plotly_chart(fig_r, use_container_width=True)
        with col2:
            fig_c = px.bar(compare, x='자치구', y='CCTV대수', color='자치구', title='CCTV 대수 비교')
            plotly_dark_layout(fig_c, height=380)
            fig_c.update_layout(showlegend=False)
            st.plotly_chart(fig_c, use_container_width=True)

        # 산점도: CCTV vs 범죄율
        st.markdown('<div class="panel-card"><div class="panel-title">🔗 CCTV 대수 vs 범죄율 관계</div></div>', unsafe_allow_html=True)
        all_compare = crime[['자치구별', col_r]].copy()
        all_cctv = cctv[['자치구','총계']].rename(columns={'자치구':'자치구별'})
        scatter = pd.merge(all_compare, all_cctv, on='자치구별', how='inner')
        scatter.columns = ['자치구','범죄율','CCTV대수']
        scatter['선택'] = scatter['자치구'].isin(selected)

        fig_s = px.scatter(
            scatter, x='CCTV대수', y='범죄율', text='자치구',
            color='선택', color_discrete_map={True: '#22d3ee', False: '#334155'},
            size=[14 if s else 8 for s in scatter['선택']],
            labels={'CCTV대수': 'CCTV 총 대수', '범죄율': f'범죄율(%)'},
        )
        fig_s.update_traces(textposition='top center', textfont=dict(size=10, color='#94a3b8'))
        plotly_dark_layout(fig_s, height=500)
        fig_s.update_layout(showlegend=False)

        # 추세선
        from numpy.polynomial.polynomial import polyfit
        x_vals = scatter['CCTV대수'].values
        y_vals = scatter['범죄율'].values
        b, m_coef = polyfit(x_vals, y_vals, 1)
        x_line = np.linspace(x_vals.min(), x_vals.max(), 100)
        y_line = b + m_coef * x_line
        fig_s.add_trace(go.Scatter(
            x=x_line, y=y_line, mode='lines',
            line=dict(color='rgba(239,68,68,0.4)', width=2, dash='dash'),
            name='추세선', showlegend=False
        ))

        st.plotly_chart(fig_s, use_container_width=True)
    else:
        st.info("비교할 자치구를 선택해주세요.")


# ============================================================
# 페이지 5: CCTV 예측 시뮬레이터 (NEW!)
# ============================================================
elif menu == "🔮 CCTV 예측 시뮬레이터":

    st.markdown('''
    <div class="panel-card">
        <div class="panel-title">🔮 CCTV 추가 설치 시 범죄 감소 예측 시뮬레이터</div>
        <div style="font-size:13px;color:#94a3b8;line-height:1.7;">
            자치구를 선택하고 CCTV 추가 설치 대수를 조절하면, 예상 범죄 감소율을 확인할 수 있습니다.<br>
            이 모델은 CCTV 설치 밀도와 범죄율 간의 상관관계를 기반으로 한 추정입니다.
        </div>
    </div>
    ''', unsafe_allow_html=True)

    sim_col1, sim_col2 = st.columns([1, 2])

    with sim_col1:
        gu_list = crime['자치구별'].tolist()
        sel_gu = st.selectbox("🏢 자치구 선택", gu_list, index=0)
        add_cctv = st.slider("📹 CCTV 추가 설치 대수", min_value=100, max_value=3000, value=500, step=100)

        # 현재 데이터
        gu_crime_rate = crime.loc[crime['자치구별']==sel_gu, col_r].values
        gu_crime_rate = float(gu_crime_rate[0]) if len(gu_crime_rate) > 0 else 0
        gu_arrest_rate = crime.loc[crime['자치구별']==sel_gu, col_a].values
        gu_arrest_rate = float(gu_arrest_rate[0]) if len(gu_arrest_rate) > 0 else 0
        gu_occur_val = occur.loc[occur['자치구별']==sel_gu, col_o].values
        gu_occur_val = int(gu_occur_val[0]) if len(gu_occur_val) > 0 else 0
        gu_cctv = cctv.loc[cctv['자치구']==sel_gu, '총계'].values
        gu_cctv = int(gu_cctv[0]) if len(gu_cctv) > 0 else 0
        gu_pop = pop.loc[pop['자치구별']==sel_gu, col_p].values if col_p in pop.columns else []
        gu_pop = int(gu_pop[0]) if len(gu_pop) > 0 else 1

        current_per1k = gu_cctv / gu_pop * 1000
        risk_label, risk_color, risk_emoji = get_risk_level(gu_crime_rate)

        st.markdown(f'''
        <div class="panel-card">
            <div class="panel-title">{risk_emoji} {sel_gu} 현재 현황</div>
            <div style="font-size:13px;line-height:2;color:#cbd5e1;">
                📹 CCTV: <b style="color:#22d3ee;">{gu_cctv:,}대</b><br>
                👥 인구: <b>{gu_pop:,}명</b><br>
                📊 범죄율: <b style="color:{risk_color};">{gu_crime_rate}%</b><br>
                🔍 검거율: <b style="color:#10b981;">{gu_arrest_rate}%</b><br>
                🚨 범죄 발생: <b style="color:#f97316;">{gu_occur_val:,}건</b><br>
                📏 천명당 CCTV: <b>{current_per1k:.1f}대</b>
            </div>
        </div>
        ''', unsafe_allow_html=True)

    with sim_col2:
        # 예측 계산
        new_total = gu_cctv + add_cctv
        new_per1k = new_total / gu_pop * 1000
        delta_per1k = new_per1k - current_per1k

        # 선형 모델: CCTV per 1k 증가 → 범죄율 감소 (데이터 상관관계 기반)
        crime_reduction = min(delta_per1k * 0.03, gu_crime_rate * 0.5)
        new_crime_rate = max(0.1, gu_crime_rate - crime_reduction)
        reduction_pct = (crime_reduction / gu_crime_rate * 100) if gu_crime_rate > 0 else 0
        est_crimes_reduced = int(gu_occur_val * (crime_reduction / gu_crime_rate)) if gu_crime_rate > 0 else 0

        st.markdown(f'''
        <div class="sim-result-card">
            <div style="font-size:16px;font-weight:700;color:#22d3ee;margin-bottom:4px;">
                {sel_gu}에 CCTV {add_cctv:,}대 추가 설치 시 예측 결과
            </div>
            <div class="sim-result-grid">
                <div class="sim-result-item">
                    <div class="sim-result-label">예상 범죄율</div>
                    <div class="sim-result-value" style="color:#10b981;">{new_crime_rate:.2f}%</div>
                    <div style="font-size:11px;color:#10b981;margin-top:4px;">▼ {crime_reduction:.2f}%p 감소</div>
                </div>
                <div class="sim-result-item">
                    <div class="sim-result-label">범죄 감소율</div>
                    <div class="sim-result-value" style="color:#10b981;">-{reduction_pct:.1f}%</div>
                    <div style="font-size:11px;color:#64748b;margin-top:4px;">현재 대비</div>
                </div>
                <div class="sim-result-item">
                    <div class="sim-result-label">예상 감소 건수</div>
                    <div class="sim-result-value" style="color:#10b981;">약 {est_crimes_reduced:,}건</div>
                    <div style="font-size:11px;color:#64748b;margin-top:4px;">{year}년 기준</div>
                </div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:14px;">
                <div class="sim-result-item">
                    <div class="sim-result-label">변경 후 CCTV 총계</div>
                    <div class="sim-result-value" style="color:#22d3ee;">{new_total:,}대</div>
                </div>
                <div class="sim-result-item">
                    <div class="sim-result-label">변경 후 천명당 CCTV</div>
                    <div class="sim-result-value" style="color:#22d3ee;">{new_per1k:.1f}대</div>
                </div>
            </div>
            <div class="sim-note">
                ※ 본 예측은 CCTV 설치 밀도와 범죄율 간의 <b>선형 회귀 추정</b>입니다.<br>
                실제 결과는 지역 특성, 경찰 배치, 인구 밀도 등 다양한 요인에 의해 달라질 수 있습니다.<br>
                <span style="color:#22d3ee;">🔬 한윤수 팀원의 ML 모델(XGBoost/Random Forest) 연동 시 더 정확한 예측이 가능합니다.</span>
            </div>
        </div>
        ''', unsafe_allow_html=True)

    # ============================================================
    # 황준연 ML 모델 결과 연동
    # ============================================================
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown('''
    <div class="panel-card">
        <div class="panel-title">🔬 ML 모델 기반 CCTV 추가 설치 우선순위 분석 (황준연)</div>
        <div style="font-size:13px;color:#94a3b8;line-height:1.7;">
            범죄율, CCTV 밀도, CCTV 1대당 범죄 부담을 종합한 우선순위 점수 기반 분류 모델 결과입니다.<br>
            가중치: 범죄율 40% + CCTV 밀도 부족도 40% + CCTV 1대당 범죄 부담 20%
        </div>
    </div>
    ''', unsafe_allow_html=True)

    # 모델 결과 CSV 로딩
    try:
        model_df = pd.read_csv('cctv_model_result_revised.csv', encoding='utf-8-sig')
        model_loaded = True
    except Exception as e:
        st.warning(f"모델 결과 파일 로딩 실패: {e}")
        model_loaded = False

    if model_loaded:
        # 모델 성능 비교 차트
        st.markdown('<div class="panel-card"><div class="panel-title">📊 모델별 정확도 비교</div></div>', unsafe_allow_html=True)
        model_perf = pd.DataFrame({
            '모델': ['Decision Tree', 'Random Forest', 'SVM', 'KNN'],
            '정확도': [0.857, 0.857, 0.714, 0.571]
        }).sort_values('정확도', ascending=False)

        fig_perf = px.bar(
            model_perf, x='모델', y='정확도',
            color='정확도', color_continuous_scale='teal',
            text=model_perf['정확도'].apply(lambda x: f'{x:.1%}'),
            labels={'정확도': '정확도 (Accuracy)', '모델': ''}
        )
        fig_perf.update_traces(textposition='outside')
        plotly_dark_layout(fig_perf, height=400)
        fig_perf.update_layout(coloraxis_showscale=False, yaxis=dict(range=[0, 1]))
        st.plotly_chart(fig_perf, use_container_width=True)

        # CCTV 추가 설치 우선순위 표
        st.markdown('<div class="panel-card"><div class="panel-title">🎯 CCTV 추가 설치 우선순위 (ML 모델 결과)</div></div>', unsafe_allow_html=True)
        priority_df = model_df.sort_values('추가설치우선순위점수', ascending=False)

        fig_priority = px.bar(
            priority_df, x='자치구', y='추가설치우선순위점수',
            color='추가설치필요도',
            color_discrete_map={'높음': '#ef4444', '보통': '#f97316', '낮음': '#10b981'},
            text=priority_df['추가설치우선순위점수'].apply(lambda x: f'{x:.1f}'),
            labels={'추가설치우선순위점수': '우선순위 점수', '자치구': '', '추가설치필요도': '설치 필요도'}
        )
        fig_priority.update_traces(textposition='outside')
        plotly_dark_layout(fig_priority, height=500)
        fig_priority.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_priority, use_container_width=True)

        # 분석 인사이트
        st.markdown('''
        <div class="panel-card">
            <div class="panel-title">💡 주요 분석 결과</div>
            <div style="font-size:13px;color:#cbd5e1;line-height:2;">
                • <b style="color:#ef4444;">중구</b>는 인구 1만 명당 범죄 수가 가장 높지만, 인구 1만 명당 CCTV 수 역시 매우 높아 CCTV 밀도 부족도가 낮게 반영되어 최상위 우선순위에서 제외됨<br>
                • <b style="color:#ef4444;">송파구</b>는 범죄율이 평균 이상이면서 CCTV 밀도가 상대적으로 낮고, CCTV 1대당 범죄 부담이 높아 추가 설치 우선순위가 가장 높게 산출됨<br>
                • 최종 선택 모델: <b style="color:#22d3ee;">Decision Tree / Random Forest</b> (정확도 85.7%)
            </div>
        </div>
        ''', unsafe_allow_html=True)

        # CCTV 밀도와 범죄율 산점도
        if '인구1만명당범죄수' in model_df.columns and '인구1만명당CCTV수' in model_df.columns:
            st.markdown('<div class="panel-card"><div class="panel-title">🔗 CCTV 밀도 vs 범죄율 관계 (ML 분석)</div></div>', unsafe_allow_html=True)

            fig_scatter = px.scatter(
                model_df, x='인구1만명당CCTV수', y='인구1만명당범죄수',
                color='추가설치필요도', text='자치구',
                color_discrete_map={'높음': '#ef4444', '보통': '#f97316', '낮음': '#10b981'},
                size='CCTV1대당범죄수',
                labels={'인구1만명당CCTV수': '인구 1만명당 CCTV 수', '인구1만명당범죄수': '인구 1만명당 범죄 수', '추가설치필요도': '설치 필요도'}
            )
            fig_scatter.update_traces(textposition='top center', textfont=dict(size=10, color='#94a3b8'))

            # 평균선
            avg_crime = model_df['인구1만명당범죄수'].mean()
            avg_cctv = model_df['인구1만명당CCTV수'].mean()
            fig_scatter.add_hline(y=avg_crime, line_dash='dash', line_color='rgba(255,255,255,0.3)', annotation_text='범죄율 평균')
            fig_scatter.add_vline(x=avg_cctv, line_dash='dash', line_color='rgba(255,255,255,0.3)', annotation_text='CCTV 밀도 평균')

            plotly_dark_layout(fig_scatter, height=550)
            st.plotly_chart(fig_scatter, use_container_width=True)

        # CCTV 추가 설치 필요도 지도
        st.markdown('<div class="panel-card"><div class="panel-title">🗺️ CCTV 추가 설치 필요도 지도 (ML 모델 결과)</div></div>', unsafe_allow_html=True)
        m_need = folium.Map(location=[37.5665, 126.9780], zoom_start=11, tiles='CartoDB dark_matter')

        folium.Choropleth(
            geo_data=seoul_geo,
            data=priority_df,
            columns=['자치구', '추가설치우선순위점수'],
            key_on='feature.properties.name',
            fill_color='YlOrRd',
            fill_opacity=0.7,
            line_opacity=0.3,
            legend_name='CCTV 추가 설치 우선순위 점수'
        ).add_to(m_need)

        need_map_data = priority_df.set_index('자치구')
        for gu, coord in gu_coords.items():
            if gu in need_map_data.index:
                row = need_map_data.loc[gu]
                score = row['추가설치우선순위점수']
                need_label = row['추가설치필요도']
                pred_label = row['예측_추가설치필요도']
                color = '#ef4444' if need_label == '높음' else '#f97316' if need_label == '보통' else '#10b981'
                folium.CircleMarker(
                    location=coord, radius=max(5, score / 5),
                    color=color, fill=True, fill_color=color, fill_opacity=0.7,
                    tooltip=f"{gu} | 점수: {score:.1f} | 필요도: {need_label} | 예측: {pred_label}"
                ).add_to(m_need)

        st_folium(m_need, width=None, height=500, returned_objects=[])

        # 전체 자치구 상세 테이블
        st.markdown("**📋 전체 자치구 CCTV 추가 설치 우선순위 상세**")
        display_cols = ['자치구', '총범죄수', 'CCTV수', '인구수', '면적',
                       '인구1만명당범죄수', '인구1만명당CCTV수', 'CCTV1대당범죄수',
                       '추가설치우선순위점수', '추가설치필요도', '예측_추가설치필요도']
        available_cols = [c for c in display_cols if c in priority_df.columns]
        display_df = priority_df[available_cols].reset_index(drop=True)
        display_df.index = display_df.index + 1
        st.dataframe(display_df, use_container_width=True)

    # ============================================================
    # 한윤수 XGBoost 모델 결과 연동
    # ============================================================
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown('''
    <div class="panel-card">
        <div class="panel-title">🤖 XGBoost 범죄 위험도 예측 모델 분석 (한윤수)</div>
        <div style="font-size:13px;color:#94a3b8;line-height:1.7;">
            야간 유동인구, 노후주택 비율, 고시원 수, 조도 지수 등 다변량 환경 요인을 반영한 XGBoost 모델 기반 예측 결과입니다.
        </div>
    </div>
    ''', unsafe_allow_html=True)

    try:
        xgb_df = pd.read_csv('Seoul_Crime_Model_Data.csv', encoding='utf-8-sig')
        # 서울 25개 자치구만 필터링
        seoul_gu_list = list(gu_coords.keys())
        xgb_seoul = xgb_df[xgb_df['자치구'].isin(seoul_gu_list)].copy()
        xgb_loaded = True
    except Exception as e:
        st.warning(f"한윤수 모델 데이터 로딩 실패: {e}")
        xgb_loaded = False

    if xgb_loaded and len(xgb_seoul) > 0:
        xgb_seoul = xgb_seoul.sort_values('예측_위험도_점수', ascending=False)

        # KPI 카드
        max_risk_gu = xgb_seoul.iloc[0]['자치구']
        max_risk_score = xgb_seoul.iloc[0]['예측_위험도_점수']
        max_priority_gu = xgb_seoul.sort_values('CCTV_설치_우선순위_점수', ascending=False).iloc[0]['자치구']
        max_priority_score = xgb_seoul.sort_values('CCTV_설치_우선순위_점수', ascending=False).iloc[0]['CCTV_설치_우선순위_점수']
        avg_risk = xgb_seoul['예측_위험도_점수'].mean()

        st.markdown(make_kpi_html([
            ('🔴 최고 위험도', max_risk_gu, f'{max_risk_score:.1f}점', 'red'),
            ('🎯 최우선 설치', max_priority_gu, f'{max_priority_score:.1f}점', 'orange'),
            ('📊 평균 위험도', f'{avg_risk:.1f}점', '서울시 25개 자치구', 'cyan'),
            ('🤖 모델', 'XGBoost', '다변량 환경 요인 반영', 'purple'),
        ]), unsafe_allow_html=True)

        # 예측 위험도 점수 차트
        st.markdown('<div class="panel-card"><div class="panel-title">🔴 자치구별 예측 위험도 점수 (XGBoost)</div></div>', unsafe_allow_html=True)
        fig_risk = px.bar(
            xgb_seoul, x='자치구', y='예측_위험도_점수',
            color='예측_위험도_점수', color_continuous_scale='YlOrRd',
            text=xgb_seoul['예측_위험도_점수'].apply(lambda x: f'{x:.1f}'),
            labels={'예측_위험도_점수': '위험도 점수', '자치구': ''}
        )
        fig_risk.update_traces(textposition='outside')
        plotly_dark_layout(fig_risk, height=450)
        fig_risk.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
        st.plotly_chart(fig_risk, use_container_width=True)

        # CCTV 설치 우선순위 점수 차트
        xgb_priority = xgb_seoul.sort_values('CCTV_설치_우선순위_점수', ascending=False)
        st.markdown('<div class="panel-card"><div class="panel-title">🎯 자치구별 CCTV 설치 우선순위 점수 (XGBoost)</div></div>', unsafe_allow_html=True)
        fig_xgb_pri = px.bar(
            xgb_priority, x='자치구', y='CCTV_설치_우선순위_점수',
            color='CCTV_설치_우선순위_점수', color_continuous_scale='Purples',
            text=xgb_priority['CCTV_설치_우선순위_점수'].apply(lambda x: f'{x:.1f}'),
            labels={'CCTV_설치_우선순위_점수': '우선순위 점수', '자치구': ''}
        )
        fig_xgb_pri.update_traces(textposition='outside')
        plotly_dark_layout(fig_xgb_pri, height=450)
        fig_xgb_pri.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
        st.plotly_chart(fig_xgb_pri, use_container_width=True)

        # 위험도 vs 우선순위 산점도
        st.markdown('<div class="panel-card"><div class="panel-title">🔗 위험도 점수 vs CCTV 설치 우선순위 (XGBoost)</div></div>', unsafe_allow_html=True)
        fig_xgb_scatter = px.scatter(
            xgb_seoul, x='예측_위험도_점수', y='CCTV_설치_우선순위_점수',
            text='자치구', size='CCTV_대수',
            color='예측_위험도_점수', color_continuous_scale='YlOrRd',
            labels={'예측_위험도_점수': '예측 위험도 점수', 'CCTV_설치_우선순위_점수': 'CCTV 설치 우선순위 점수'}
        )
        fig_xgb_scatter.update_traces(textposition='top center', textfont=dict(size=10, color='#94a3b8'))
        plotly_dark_layout(fig_xgb_scatter, height=500)
        fig_xgb_scatter.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig_xgb_scatter, use_container_width=True)

        # XGBoost 위험도 지도
        st.markdown('<div class="panel-card"><div class="panel-title">🗺️ XGBoost 예측 위험도 지도</div></div>', unsafe_allow_html=True)
        m_xgb = folium.Map(location=[37.5665, 126.9780], zoom_start=11, tiles='CartoDB dark_matter')

        folium.Choropleth(
            geo_data=seoul_geo,
            data=xgb_seoul,
            columns=['자치구', '예측_위험도_점수'],
            key_on='feature.properties.name',
            fill_color='YlOrRd',
            fill_opacity=0.7,
            line_opacity=0.3,
            legend_name='XGBoost 예측 위험도 점수'
        ).add_to(m_xgb)

        xgb_map = xgb_seoul.drop_duplicates(subset='자치구', keep='first').set_index('자치구')
        for gu, coord in gu_coords.items():
            if gu in xgb_map.index:
                row = xgb_map.loc[gu]
                risk = float(row['예측_위험도_점수'])
                pri = float(row['CCTV_설치_우선순위_점수'])
                cctv_cnt = float(row['CCTV_대수'])
                color = '#ef4444' if risk >= 60 else '#f97316' if risk >= 30 else '#10b981'
                folium.CircleMarker(
                    location=coord, radius=max(5, risk / 7),
                    color=color, fill=True, fill_color=color, fill_opacity=0.7,
                    tooltip=f"{gu} | 위험도: {risk:.1f} | 우선순위: {pri:.1f} | CCTV: {cctv_cnt:.0f}대"
                ).add_to(m_xgb)

        st_folium(m_xgb, width=None, height=500, returned_objects=[])

        # 상세 데이터 테이블
        st.markdown("**📋 XGBoost 모델 전체 분석 결과 (서울 25개 자치구)**")
        xgb_display = xgb_seoul[['자치구', '야간_유동인구', 'CCTV_대수', '발생_10만명당',
                                  '예측_위험도_점수', 'CCTV_설치_우선순위_점수', '조도_지수']].reset_index(drop=True)
        xgb_display.index = xgb_display.index + 1
        st.dataframe(xgb_display, use_container_width=True)

# ============================================================
# 페이지 6: 범죄 유형별 분석
# ============================================================
elif menu == "🔍 범죄 유형별 분석":
    st.markdown('<div class="page-header"><span class="page-icon">🔍</span>5대 범죄 유형별 분석 (2024년)</div>', unsafe_allow_html=True)

    # crime_seoul.csv에서 유형별 데이터 파싱
    try:
        crime_type_raw = pd.read_csv('crime_seoul.csv', encoding='utf-8-sig')
        crime_type_data = crime_type_raw.iloc[3:].copy()
        crime_type_df = pd.DataFrame({
            '자치구': crime_type_data.iloc[:, 1].astype(str).str.strip(),
            '총범죄_발생': pd.to_numeric(crime_type_data.iloc[:, 2].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '총범죄_검거': pd.to_numeric(crime_type_data.iloc[:, 3].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '살인_발생': pd.to_numeric(crime_type_data.iloc[:, 4].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '살인_검거': pd.to_numeric(crime_type_data.iloc[:, 5].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '강도_발생': pd.to_numeric(crime_type_data.iloc[:, 6].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '강도_검거': pd.to_numeric(crime_type_data.iloc[:, 7].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '강간강제추행_발생': pd.to_numeric(crime_type_data.iloc[:, 8].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '강간강제추행_검거': pd.to_numeric(crime_type_data.iloc[:, 9].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '절도_발생': pd.to_numeric(crime_type_data.iloc[:, 10].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '절도_검거': pd.to_numeric(crime_type_data.iloc[:, 11].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '폭력_발생': pd.to_numeric(crime_type_data.iloc[:, 12].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
            '폭력_검거': pd.to_numeric(crime_type_data.iloc[:, 13].astype(str).str.replace(',', '').str.replace('-', '0'), errors='coerce'),
        })
        crime_type_df = crime_type_df[crime_type_df['자치구'] != '소계'].dropna()
        crime_type_loaded = True
    except Exception as e:
        st.warning(f"범죄 유형별 데이터 로딩 실패: {e}")
        crime_type_loaded = False

    if crime_type_loaded:
        # 서울시 전체 요약 KPI
        total_row = crime_type_df.sum(numeric_only=True)
        st.markdown(make_kpi_html([
            ('🔪 살인', f"{int(total_row['살인_발생'])}건", f"검거 {int(total_row['살인_검거'])}건", 'red'),
            ('🔫 강도', f"{int(total_row['강도_발생'])}건", f"검거 {int(total_row['강도_검거'])}건", 'orange'),
            ('⚠️ 강간·강제추행', f"{int(total_row['강간강제추행_발생'])}건", f"검거 {int(total_row['강간강제추행_검거'])}건", 'purple'),
            ('🛍️ 절도', f"{int(total_row['절도_발생'])}건", f"검거 {int(total_row['절도_검거'])}건", 'cyan'),
        ]), unsafe_allow_html=True)

        # 유형 선택
        crime_types = ['살인', '강도', '강간강제추행', '절도', '폭력']
        selected_type = st.selectbox("📌 분석할 범죄 유형 선택", crime_types)

        col_occur = f'{selected_type}_발생'
        col_arrest = f'{selected_type}_검거'

        sorted_df = crime_type_df.sort_values(col_occur, ascending=False)

        # 발생 건수 바차트
        st.markdown(f'<div class="panel-card"><div class="panel-title">📊 자치구별 {selected_type} 발생 건수</div></div>', unsafe_allow_html=True)
        fig_type = px.bar(
            sorted_df, x='자치구', y=col_occur,
            color=col_occur, color_continuous_scale='YlOrRd',
            text=sorted_df[col_occur].apply(lambda x: f'{int(x)}'),
            labels={col_occur: '발생 건수', '자치구': ''}
        )
        fig_type.update_traces(textposition='outside')
        plotly_dark_layout(fig_type, height=450)
        fig_type.update_layout(xaxis_tickangle=-45, coloraxis_showscale=False)
        st.plotly_chart(fig_type, use_container_width=True)

        # 발생 vs 검거 비교
        st.markdown(f'<div class="panel-card"><div class="panel-title">⚖️ {selected_type} 발생 vs 검거 비교</div></div>', unsafe_allow_html=True)
        fig_compare = go.Figure()
        fig_compare.add_trace(go.Bar(name='발생', x=sorted_df['자치구'], y=sorted_df[col_occur], marker_color='#ef4444'))
        fig_compare.add_trace(go.Bar(name='검거', x=sorted_df['자치구'], y=sorted_df[col_arrest], marker_color='#22d3ee'))
        fig_compare.update_layout(barmode='group')
        plotly_dark_layout(fig_compare, height=450)
        fig_compare.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_compare, use_container_width=True)

        # 검거율 차트
        st.markdown(f'<div class="panel-card"><div class="panel-title">🎯 자치구별 {selected_type} 검거율</div></div>', unsafe_allow_html=True)
        arrest_df = sorted_df.copy()
        arrest_df['검거율'] = (arrest_df[col_arrest] / arrest_df[col_occur].replace(0, 1) * 100).round(1)
        arrest_df = arrest_df.sort_values('검거율', ascending=True)

        fig_arrest = px.bar(
            arrest_df, x='검거율', y='자치구', orientation='h',
            color='검거율', color_continuous_scale='Teal',
            text=arrest_df['검거율'].apply(lambda x: f'{x:.1f}%'),
            labels={'검거율': '검거율 (%)', '자치구': ''}
        )
        fig_arrest.update_traces(textposition='outside')
        plotly_dark_layout(fig_arrest, height=550)
        fig_arrest.update_layout(yaxis=dict(autorange='reversed'), coloraxis_showscale=False)
        st.plotly_chart(fig_arrest, use_container_width=True)

        # 5대 범죄 유형 구성 비율 (전체)
        st.markdown('<div class="panel-card"><div class="panel-title">🥧 서울시 5대 범죄 유형 구성 비율</div></div>', unsafe_allow_html=True)
        pie_data = pd.DataFrame({
            '유형': ['살인', '강도', '강간·강제추행', '절도', '폭력'],
            '발생건수': [int(total_row['살인_발생']), int(total_row['강도_발생']),
                      int(total_row['강간강제추행_발생']), int(total_row['절도_발생']), int(total_row['폭력_발생'])]
        })
        fig_pie = px.pie(pie_data, values='발생건수', names='유형',
                         color_discrete_sequence=['#ef4444', '#f97316', '#a855f7', '#22d3ee', '#eab308'])
        plotly_dark_layout(fig_pie, height=400)
        st.plotly_chart(fig_pie, use_container_width=True)

        # 자치구별 유형별 히트맵
        st.markdown('<div class="panel-card"><div class="panel-title">🔥 자치구별 5대 범죄 발생 히트맵</div></div>', unsafe_allow_html=True)
        heatmap_df = crime_type_df.set_index('자치구')[['살인_발생', '강도_발생', '강간강제추행_발생', '절도_발생', '폭력_발생']]
        heatmap_df.columns = ['살인', '강도', '강간·강제추행', '절도', '폭력']

        fig_heat = px.imshow(
            heatmap_df.values,
            labels=dict(x="범죄 유형", y="자치구", color="발생 건수"),
            x=heatmap_df.columns.tolist(),
            y=heatmap_df.index.tolist(),
            color_continuous_scale='YlOrRd',
            aspect='auto'
        )
        plotly_dark_layout(fig_heat, height=700)
        st.plotly_chart(fig_heat, use_container_width=True)

        # 상세 테이블
        st.markdown("**📋 전체 자치구 5대 범죄 유형별 상세 데이터**")
        table_df = crime_type_df[['자치구', '살인_발생', '강도_발생', '강간강제추행_발생', '절도_발생', '폭력_발생', '총범죄_발생']].sort_values('총범죄_발생', ascending=False).reset_index(drop=True)
        table_df.index = table_df.index + 1
        st.dataframe(table_df, use_container_width=True)

# ============================================================
# 페이지 7: 범죄 추이 예측
# ============================================================
elif menu == "📉 범죄 추이 예측":
    st.markdown('<div class="page-header"><span class="page-icon">📉</span>범죄 추이 예측 — 연도별 트렌드 & 미래 전망</div>', unsafe_allow_html=True)

    # 5개년 데이터 로딩
    try:
        for enc in ['cp949', 'euc-kr', 'utf-8', 'utf-8-sig']:
            try:
                trend_raw = pd.read_csv('자치구별 범죄율 검거율 5개년.csv', encoding=enc)
                break
            except:
                continue

        # 범죄율 파싱
        years_5 = [2019, 2020, 2021, 2022, 2023]
        trend_list = []
        for _, row in trend_raw.iloc[1:].iterrows():
            gu = str(row.iloc[0]).strip()
            if not gu or gu == 'nan':
                continue
            for i, yr in enumerate(years_5):
                cr_val = pd.to_numeric(str(row.iloc[1 + i * 2]).replace(',', ''), errors='coerce')
                ar_val = pd.to_numeric(str(row.iloc[2 + i * 2]).replace(',', ''), errors='coerce')
                trend_list.append({'자치구': gu, '연도': yr, '범죄율': cr_val, '검거율': ar_val})

        # 2024년 데이터 추가 (crime에서)
        if col_r in crime.columns and col_a in crime.columns:
            for _, row in crime.iterrows():
                gu = row['자치구별']
                cr = row[col_r] if col_r in crime.columns else None
                ar = row[col_a] if col_a in crime.columns else None
                if pd.notna(cr):
                    trend_list.append({'자치구': gu, '연도': 2024, '범죄율': cr, '검거율': ar})

        trend_df = pd.DataFrame(trend_list).dropna(subset=['범죄율'])
        trend_loaded = True
    except Exception as e:
        st.warning(f"추이 데이터 로딩 실패: {e}")
        trend_loaded = False

    if trend_loaded:
        available_years = sorted(trend_df['연도'].unique())
        all_gus = sorted(trend_df['자치구'].unique())

        # 자치구 선택
        selected_gus = st.multiselect("📍 분석할 자치구 선택 (최대 5개)", all_gus, default=['강남구', '송파구', '중구'], max_selections=5)

        if selected_gus:
            filtered = trend_df[trend_df['자치구'].isin(selected_gus)]

            # 범죄율 추이 차트
            st.markdown('<div class="panel-card"><div class="panel-title">📈 자치구별 범죄율 연도별 추이</div></div>', unsafe_allow_html=True)
            fig_trend = px.line(
                filtered, x='연도', y='범죄율', color='자치구',
                markers=True, labels={'범죄율': '범죄율 (%)', '연도': '연도'}
            )
            plotly_dark_layout(fig_trend, height=450)
            fig_trend.update_layout(xaxis=dict(dtick=1))
            st.plotly_chart(fig_trend, use_container_width=True)

            # 미래 예측 (선형 회귀)
            st.markdown('<div class="panel-card"><div class="panel-title">🔮 범죄율 추세 예측 (2025~2026년)</div></div>', unsafe_allow_html=True)

            fig_pred = go.Figure()
            pred_years = [2025, 2026]

            for gu in selected_gus:
                gu_data = filtered[filtered['자치구'] == gu].sort_values('연도')
                x = gu_data['연도'].values
                y = gu_data['범죄율'].values

                if len(x) >= 2:
                    # 선형 회귀
                    coeffs = np.polyfit(x, y, 1)
                    poly = np.poly1d(coeffs)

                    # 실제 데이터
                    fig_pred.add_trace(go.Scatter(
                        x=x, y=y, mode='lines+markers', name=f'{gu} (실측)',
                        line=dict(width=2)
                    ))

                    # 예측
                    all_x = list(x) + pred_years
                    pred_y = [poly(yr) for yr in pred_years]
                    fig_pred.add_trace(go.Scatter(
                        x=pred_years, y=[max(0, p) for p in pred_y],
                        mode='lines+markers', name=f'{gu} (예측)',
                        line=dict(dash='dash', width=2),
                        marker=dict(symbol='diamond', size=10)
                    ))

            plotly_dark_layout(fig_pred, height=500)
            fig_pred.update_layout(xaxis=dict(dtick=1, title='연도'), yaxis_title='범죄율 (%)')
            st.plotly_chart(fig_pred, use_container_width=True)

            # 예측 결과 테이블
            st.markdown('<div class="panel-card"><div class="panel-title">📋 2025~2026 범죄율 예측 결과</div></div>', unsafe_allow_html=True)
            pred_table = []
            for gu in selected_gus:
                gu_data = filtered[filtered['자치구'] == gu].sort_values('연도')
                x = gu_data['연도'].values
                y = gu_data['범죄율'].values
                if len(x) >= 2:
                    coeffs = np.polyfit(x, y, 1)
                    poly = np.poly1d(coeffs)
                    cr_2024 = gu_data[gu_data['연도'] == gu_data['연도'].max()]['범죄율'].values[0]
                    cr_2025 = max(0, poly(2025))
                    cr_2026 = max(0, poly(2026))
                    change = cr_2026 - cr_2024
                    pred_table.append({
                        '자치구': gu,
                        f'{int(gu_data["연도"].max())}년 범죄율': f'{cr_2024:.2f}%',
                        '2025년 예측': f'{cr_2025:.2f}%',
                        '2026년 예측': f'{cr_2026:.2f}%',
                        '변화 추세': '📈 증가' if change > 0.05 else '📉 감소' if change < -0.05 else '➡️ 유지'
                    })
            pred_result = pd.DataFrame(pred_table)
            pred_result.index = pred_result.index + 1
            st.dataframe(pred_result, use_container_width=True)

            # 검거율 추이
            if '검거율' in filtered.columns:
                st.markdown('<div class="panel-card"><div class="panel-title">🎯 자치구별 검거율 연도별 추이</div></div>', unsafe_allow_html=True)
                fig_arrest_trend = px.line(
                    filtered, x='연도', y='검거율', color='자치구',
                    markers=True, labels={'검거율': '검거율 (%)', '연도': '연도'}
                )
                plotly_dark_layout(fig_arrest_trend, height=400)
                fig_arrest_trend.update_layout(xaxis=dict(dtick=1))
                st.plotly_chart(fig_arrest_trend, use_container_width=True)

            # 서울 전체 평균 추이
            st.markdown('<div class="panel-card"><div class="panel-title">🏙️ 서울시 전체 평균 범죄율 추이</div></div>', unsafe_allow_html=True)
            avg_trend = trend_df.groupby('연도').agg({'범죄율': 'mean', '검거율': 'mean'}).reset_index()
            fig_avg = go.Figure()
            fig_avg.add_trace(go.Scatter(x=avg_trend['연도'], y=avg_trend['범죄율'], mode='lines+markers', name='평균 범죄율', line=dict(color='#ef4444', width=3)))
            fig_avg.add_trace(go.Scatter(x=avg_trend['연도'], y=avg_trend['검거율'], mode='lines+markers', name='평균 검거율', line=dict(color='#22d3ee', width=3), yaxis='y2'))
            plotly_dark_layout(fig_avg, height=400)
            fig_avg.update_layout(
                xaxis=dict(dtick=1, title='연도'),
                yaxis=dict(title='범죄율 (%)'),
                yaxis2=dict(title='검거율 (%)', overlaying='y', side='right')
            )
            st.plotly_chart(fig_avg, use_container_width=True)
        else:
            st.info("분석할 자치구를 선택하세요.")

# ============================================================
# 페이지 8: AI 심층 분석 (한윤수 XGBoost + Gemini 질의응답)
# ============================================================
elif menu == "🧠 AI 심층 분석":
    st.markdown('<div class="page-header"><span class="page-icon">🧠</span>AI 심층 분석 — 자치구별 환경 요인 분석 & 정책 시뮬레이션</div>', unsafe_allow_html=True)

    # AI 엔진 로딩
    @st.cache_resource
    def load_ai_engine():
        data_path = 'Seoul_Crime_Model_Data.csv'
        model_path = 'xgb_crime_model.json'
        try:
            ai_df = pd.read_csv(data_path)
            ai_df = ai_df[ai_df['발생_10만명당'] > 0].copy()
            ai_features = [
                '야간_유동인구', '버스노선_총합', '노후주택_비율', '고시원_수',
                'CCTV_대수', '조도_지수', '자연감시_시너지_지수', '사각지대_취약_지수', '군중밀집_위험_지수'
            ]
            ai_scaler = MinMaxScaler()
            ai_scaler.fit(ai_df[ai_features])
            ai_model = xgb.XGBRegressor()
            ai_model.load_model(model_path)
            return ai_df, ai_scaler, ai_model, ai_features
        except Exception as e:
            st.error(f"AI 엔진 로딩 실패: {e}")
            return pd.DataFrame(), None, None, []

    ai_df, ai_scaler, ai_model, ai_features = load_ai_engine()

    if not ai_df.empty and ai_scaler is not None:
        # 자치구 선택
        seoul_gu_list_ai = sorted([g for g in gu_coords.keys() if g in ai_df['자치구'].values])
        selected_gu = st.selectbox("🔍 분석할 자치구를 선택하세요", seoul_gu_list_ai)
        gu_data = ai_df[ai_df['자치구'] == selected_gu].iloc[0]

        # 분석 모드 선택
        ai_mode = st.radio("분석 모드", ["🧠 환경 요인 분석", "🎯 CCTV 설치 우선순위", "📊 심층 진단 리포트", "🔮 미래 정책 시뮬레이션"], horizontal=True)

        if ai_mode == "🧠 환경 요인 분석":
            st.markdown(f'<div class="panel-card"><div class="panel-title">🧠 AI 딥다이브: {selected_gu} 환경 요인 분석</div></div>', unsafe_allow_html=True)
            col1, col2 = st.columns([1, 1.5])
            with col1:
                st.metric("10만명당 범죄 발생 수", f"{gu_data['발생_10만명당']:.1f}건")
                st.markdown(f'''
                <div class="panel-card" style="border-left: 3px solid #22d3ee;">
                    <b>📍 {selected_gu} 주요 지표 현황:</b><br><br>
                    • 야간 유동인구: <b>{int(gu_data['야간_유동인구']):,}</b>명<br>
                    • 노후주택 비율: <b>{gu_data['노후주택_비율']:.1f}</b>%<br>
                    • CCTV 대수: <b>{int(gu_data['CCTV_대수'])}</b>대<br>
                    • 조도 지수: <b>{gu_data['조도_지수']:.1f}</b>점
                </div>
                ''', unsafe_allow_html=True)
            with col2:
                radar_data = pd.DataFrame({
                    '지표': ai_features,
                    '현재값_상대비율': [(gu_data[f] / ai_df[f].max() * 100) if ai_df[f].max() > 0 else 0 for f in ai_features]
                }).sort_values('현재값_상대비율', ascending=True)
                fig_local = px.bar(radar_data, x='현재값_상대비율', y='지표', orientation='h',
                                   title=f"[{selected_gu}] 타 자치구 대비 위험 요인 수준 (%)")
                fig_local.update_traces(marker_color='#e84393')
                plotly_dark_layout(fig_local, height=400)
                fig_local.update_layout(xaxis_title="최대치 대비 비율(%)", yaxis_title="")
                st.plotly_chart(fig_local, use_container_width=True)

        elif ai_mode == "🎯 CCTV 설치 우선순위":
            st.markdown(f'<div class="panel-card"><div class="panel-title">🎯 {selected_gu} CCTV 설치 우선순위 리포트</div></div>', unsafe_allow_html=True)
            priority_ai = ai_df[ai_df['자치구'].isin(seoul_gu_list_ai)].sort_values('CCTV_설치_우선순위_점수', ascending=False)
            gu_rank = priority_ai.reset_index(drop=True)[priority_ai.reset_index(drop=True)['자치구'] == selected_gu].index[0] + 1

            c1, c2, c3 = st.columns(3)
            with c1: st.error(f"### 🚨 긴급 설치 1순위\n{priority_ai.iloc[0]['자치구']}")
            with c2: st.info(f"### 📍 현재 선택 지역\n**{selected_gu}** (전체 {len(priority_ai)}개 중 **{gu_rank}위**)")
            with c3: st.success(f"### ✅ 유지/관리 구역\n{priority_ai.iloc[-1]['자치구']}")

            colors = ['#ef4444' if gu == selected_gu else '#a29bfe' for gu in priority_ai['자치구']]
            fig_pri_ai = go.Figure(data=[go.Bar(x=priority_ai['자치구'], y=priority_ai['CCTV_설치_우선순위_점수'], marker_color=colors)])
            fig_pri_ai.update_layout(title=f"서울시 전체 CCTV 설치 우선순위 점수 ({selected_gu} 강조)")
            plotly_dark_layout(fig_pri_ai, height=450)
            fig_pri_ai.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_pri_ai, use_container_width=True)

        elif ai_mode == "📊 심층 진단 리포트":
            st.markdown(f'<div class="panel-card"><div class="panel-title">📊 {selected_gu} 범죄 및 치안 인프라 심층 리포트</div></div>', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("야간 유동인구", f"{int(gu_data['야간_유동인구']/10000)}만 명")
            with c2: st.metric("노후주택 비율", f"{gu_data['노후주택_비율']:.1f}%")
            with c3: st.metric("조도 지수", f"{gu_data['조도_지수']:.1f}점")
            with c4: st.metric("고시원 수", f"{int(gu_data['고시원_수'])}개")

            col1, col2 = st.columns(2)
            with col1:
                radar_cols = ['노후주택_비율', '야간_유동인구', '사각지대_취약_지수', '군중밀집_위험_지수', '고시원_수']
                radar_vals = [(gu_data[c] / ai_df[c].max()) * 100 if ai_df[c].max() > 0 else 0 for c in radar_cols]
                fig_radar = go.Figure(go.Scatterpolar(
                    r=radar_vals, theta=['노후주택', '야간인구', '사각지대', '군중밀집', '고시원'],
                    fill='toself', line_color='#22d3ee'
                ))
                fig_radar.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0, 100]), bgcolor='rgba(0,0,0,0)'),
                    showlegend=False, title=f"{selected_gu} 환경 취약점 분석", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_radar, use_container_width=True)
            with col2:
                st.markdown(f'''
                <div class="panel-card" style="border-left: 3px solid #f97316;">
                    <b>[{selected_gu} AI 진단 리포트]</b><br><br>
                    <b>🔍 CPTED 분석 결과:</b><br>
                    {selected_gu}의 사각지대 취약 지수는 <b>{gu_data['사각지대_취약_지수']:.2f}</b>이며,
                    자연감시 시너지 지수는 <b>{gu_data['자연감시_시너지_지수']:.1f}</b>입니다.<br><br>
                    <b>📍 AI 권장 액션 플랜:</b><br>
                    1. 야간 유동인구가 몰리는 22시~04시 시간대 집중 순찰<br>
                    2. 노후주택 밀집 구역 중심의 스마트 보안등 및 CCTV 우선 확충
                </div>
                ''', unsafe_allow_html=True)

        elif ai_mode == "🔮 미래 정책 시뮬레이션":
            st.markdown(f'<div class="panel-card"><div class="panel-title">🔮 {selected_gu} 미래 치안 예측 시뮬레이터 (XGBoost AI 연동)</div></div>', unsafe_allow_html=True)
            st.markdown("XGBoost 모델이 정책 변화를 실시간으로 받아 CPTED 파생 변수를 재계산하고, 미래의 범죄 발생 건수를 예측합니다.")

            orig_cctv = gu_data['CCTV_대수']
            orig_illum = gu_data['조도_지수']
            orig_pop = gu_data['야간_유동인구']

            st.sidebar.header("⚙️ 예방 정책 파라미터")
            st.sidebar.write(f"현재 CCTV: {int(orig_cctv)}대")
            cctv_input = st.sidebar.slider("CCTV 추가 설치 (대)", 0, 500, 0, 10)
            st.sidebar.write(f"현재 조도 지수: {orig_illum:.1f}점")
            illum_input = st.sidebar.slider("가로등 조도 개선 (+점수)", 0.0, 50.0, 0.0, 1.0)
            st.sidebar.write(f"현재 야간 인구: {int(orig_pop/10000)}만 명")
            pop_input = st.sidebar.slider("야간 유동인구 분산/감소율 (%)", 0.0, 50.0, 0.0, 1.0)

            if st.sidebar.button("🚀 시뮬레이션 실행", type="primary"):
                with st.spinner("AI가 새로운 정책 조건에서 예측 중입니다..."):
                    new_cctv = orig_cctv + cctv_input
                    new_illum = min(100.0, orig_illum + illum_input)
                    new_pop = orig_pop * (1 - pop_input / 100.0)

                    epsilon = 1e-5
                    threshold_pop = 500000
                    new_synergy = np.log1p(new_pop) * new_illum
                    new_blind = new_pop / ((new_cctv + epsilon) * (new_illum + epsilon))
                    new_crowd = (new_pop - threshold_pop) ** 2

                    orig_row = gu_data[ai_features].to_frame().T
                    new_row = orig_row.copy()
                    new_row['CCTV_대수'] = new_cctv
                    new_row['조도_지수'] = new_illum
                    new_row['야간_유동인구'] = new_pop
                    new_row['자연감시_시너지_지수'] = new_synergy
                    new_row['사각지대_취약_지수'] = new_blind
                    new_row['군중밀집_위험_지수'] = new_crowd

                    pred_before = max(0, ai_model.predict(ai_scaler.transform(orig_row))[0])
                    pred_after = max(0, ai_model.predict(ai_scaler.transform(new_row))[0])
                    improvement = pred_before - pred_after

                    c1, c2, c3 = st.columns(3)
                    with c1: st.metric("정책 전 예측 발생건수", f"{pred_before:.1f}건")
                    with c2: st.metric("정책 후 예측 발생건수", f"{pred_after:.1f}건", delta=f"{-improvement:.1f}건", delta_color="inverse")
                    with c3:
                        improve_rate = (improvement / pred_before * 100) if pred_before > 0 else 0
                        st.metric("범죄 억제율", f"{improve_rate:.1f}%")
            else:
                st.info("👈 좌측 사이드바에서 정책 파라미터를 조절하고 '시뮬레이션 실행' 버튼을 눌러주세요.")

        # AI 치안 정책 보좌관 (질의응답) — 모든 모드에서 표시
        st.markdown('<hr class="divider">', unsafe_allow_html=True)
        st.markdown('<div class="panel-card"><div class="panel-title">💬 AI 치안 정책 보좌관 (질의응답)</div><div style="font-size:13px;color:#94a3b8;">현재 보고 계신 자치구의 데이터나 치안 관련 궁금한 점을 질문해보세요.</div></div>', unsafe_allow_html=True)

        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = []

        for message in st.session_state.chat_messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        if prompt := st.chat_input(f"{selected_gu} 치안 데이터에 대해 무엇이든 물어보세요..."):
            st.session_state.chat_messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("AI가 분석 중입니다..."):
                    try:
                        client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                        context = f"현재 분석 중인 지역은 {selected_gu}입니다. 해당 지역의 야간 유동인구는 {gu_data['야간_유동인구']}명, CCTV는 {gu_data['CCTV_대수']}대, 노후주택 비율은 {gu_data['노후주택_비율']}%, 조도 지수는 {gu_data['조도_지수']}점입니다. 이 정보를 바탕으로 대답해 주세요. 질문: {prompt}"
                        response = client.models.generate_content(model="gemini-2.5-flash", contents=context)
                        st.markdown(response.text)
                        st.session_state.chat_messages.append({"role": "assistant", "content": response.text})
                    except Exception as e:
                        st.error(f"AI 응답 생성 실패: {e}")
    else:
        st.warning("AI 엔진 로딩에 실패했습니다. Seoul_Crime_Model_Data.csv와 xgb_crime_model.json 파일을 확인해주세요.")