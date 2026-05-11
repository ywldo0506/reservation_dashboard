import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="단체예약 현황 대시보드", page_icon="📋", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.metric-card { background:white; border-radius:12px; padding:20px 24px; box-shadow:0 2px 8px rgba(0,0,0,0.07); border-left:4px solid #2E7D32; }
.metric-card.red { border-left-color:#B71C1C; }
.metric-card.blue { border-left-color:#1565C0; }
.metric-card.orange { border-left-color:#E65100; }
.metric-label { font-size:13px; color:#666; margin-bottom:4px; }
.metric-value { font-size:24px; font-weight:700; color:#1a1a1a; }
.metric-sub { font-size:12px; color:#999; margin-top:2px; }
</style>
""", unsafe_allow_html=True)

SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ9DoJn7KHOQI44dFBskzlmZI8Xi1_tixVWJhIM-dZezbVkM8C5WpCEDXrrLS6QEHd28GCy1rjSxNme/pub?gid=1866516420&single=true&output=csv"

@st.cache_data(ttl=300)
def load_data():
    try:
        raw = pd.read_csv(SHEET_URL, skiprows=1, header=0)
    except Exception as e:
        st.error(f"데이터 로드 실패: {e}")
        return pd.DataFrame()

    raw.columns = [str(c).strip() for c in raw.columns]
    col_map = {}
    for c in raw.columns:
        if '인입' in c: col_map[c] = '인입일'
        elif '매장' in c: col_map[c] = '예약매장'
        elif '구분' in c: col_map[c] = '구분'
        elif '성인' in c: col_map[c] = '성인'
        elif '초등' in c: col_map[c] = '초등'
        elif '미취학' in c: col_map[c] = '미취학'
        elif '접수' in c: col_map[c] = '접수채널'
        elif '체결 여부' in c or '체결여부' in c: col_map[c] = '체결여부'
    raw = raw.rename(columns=col_map)

    for col in ['인입일','예약매장','구분','성인','초등','미취학','체결여부']:
        if col not in raw.columns:
            raw[col] = None

    raw = raw.dropna(subset=['인입일'])
    raw = raw[raw['예약매장'].notna() & (raw['예약매장'].astype(str).str.strip() != '') & (raw['예약매장'].astype(str).str.strip() != 'nan')]
    raw['예약매장'] = raw['예약매장'].astype(str).str.strip().apply(lambda x: re.sub(r'\s+', ' ', x))
    raw['예약매장'] = raw['예약매장'].replace({'퀸즈 여의도한강공원점':'퀸즈 여의도 한강공원점','퀸즈 구의 이스트폴점':'퀸즈 구의이스트폴점','퀸즈 천안 펜타포트점':'퀸즈 천안펜타포트점'})

    for col in ['성인','초등','미취학']:
        raw[col] = pd.to_numeric(raw[col], errors='coerce').fillna(0).astype(int)

    raw['총인원'] = raw['성인'] + raw['초등'] + raw['미취학']
    raw['구분'] = raw['구분'].replace({'런지':'런치','평일':'런치','공휴일':'주말','무응':'기타','미정':'기타'})
    raw['구분'] = raw['구분'].fillna('기타')
    raw['체결구분'] = raw['체결여부'].apply(lambda x: '체결' if pd.isna(x) or str(x).strip() in ['','보류'] else '불가')
    raw['예상금액'] = raw['성인']*25900 + raw['초등']*12900 + raw['미취학']*7900
    return raw

col_title, col_refresh = st.columns([6, 1])
with col_title:
    st.markdown("## 📋 단체예약 현황 대시보드")
with col_refresh:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 새로고침"):
        st.cache_data.clear()
        st.rerun()

df = load_data()
if df.empty:
    st.warning("데이터를 불러오지 못했어요.")
    st.stop()

st.markdown("---")
try:
    df['인입일_dt'] = pd.to_datetime(df['인입일'], errors='coerce')
    min_date = df['인입일_dt'].min().date()
    max_date = df['인입일_dt'].max().date()
    col_f1, col_f2, col_f3 = st.columns([2, 2, 4])
    with col_f1:
        start_date = st.date_input("시작일", value=min_date, min_value=min_date, max_value=max_date)
    with col_f2:
        end_date = st.date_input("종료일", value=max_date, min_value=min_date, max_value=max_date)
    df = df[(df['인입일_dt'].dt.date >= start_date) & (df['인입일_dt'].dt.date <= end_date)]
except:
    pass

st.caption(f"총 {len(df)}건 기준 · 구글 시트 실시간 연동 (5분 자동갱신)")

체결df = df[df['체결구분'] == '체결']
불가df = df[df['체결구분'] == '불가']
total = len(df)
체결수 = len(체결df)
불가수 = len(불가df)
체결률 = round(체결수/total*100, 1) if total > 0 else 0
체결금액 = 체결df['예상금액'].sum()
손실금액 = 불가df['예상금액'].sum()

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.markdown(f'<div class="metric-card blue"><div class="metric-label">전체 예약</div><div class="metric-value">{total:,}건</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card"><div class="metric-label">체결</div><div class="metric-value">{체결수:,}건</div><div class="metric-sub">체결률 {체결률}%</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-card red"><div class="metric-label">불가</div><div class="metric-value">{불가수:,}건</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="metric-card orange"><div class="metric-label">체결 예상금액</div><div class="metric-value" style="font-size:18px">{체결금액:,.0f}원</div></div>', unsafe_allow_html=True)
with c5:
    st.markdown(f'<div class="metric-card red"><div class="metric-label">손실 금액</div><div class="metric-value" style="font-size:18px">{손실금액:,.0f}원</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("#### 📊 구분별 예약 현황")
구분p = df.groupby('구분').agg(건수=('예약매장','count'), 체결=('체결구분', lambda x:(x=='체결').sum())).reset_index()
order = ['런치','주말','디너','기타']
구분p['ord'] = 구분p['구분'].apply(lambda x: order.index(x) if x in order else 99)
구분p = 구분p.sort_values('ord').drop(columns='ord')
st.bar_chart(구분p.set_index('구분')[['건수','체결']], use_container_width=True, height=200)

st.markdown("#### ① 매장별 예약 현황 (체결 기준 인원·금액)")
col_search, col_filter1, col_filter2 = st.columns([3, 1, 1])
with col_search:
    search = st.text_input("🔍 매장명 검색", placeholder="매장명 입력...")
with col_filter1:
    show_loss = st.checkbox("손실 있는 매장만")
with col_filter2:
    sort_col = st.selectbox("정렬 기준", ["총건수","체결률","체결금액","손실금액"], index=0)

p = df.groupby('예약매장').agg(총건수=('예약매장','count'), 체결=('체결구분', lambda x:(x=='체결').sum()), 불가=('체결구분', lambda x:(x=='불가').sum())).reset_index()
cp = 체결df.groupby('예약매장').agg(성인=('성인','sum'), 초등=('초등','sum'), 미취학=('미취학','sum'), 총인원=('총인원','sum'), 체결금액=('예상금액','sum')).reset_index()
lp = 불가df.groupby('예약매장').agg(손실금액=('예상금액','sum')).reset_index()
result = p.merge(cp, on='예약매장', how='left').merge(lp, on='예약매장', how='left').fillna(0)
for col in ['성인','초등','미취학','총인원','체결금액','손실금액']:
    result[col] = result[col].astype(int)
result['체결률(%)'] = (result['체결']/result['총건수']*100).round(1)
result = result.rename(columns={'성인':'체결 성인','초등':'체결 초등','미취학':'체결 미취학','총인원':'체결 총인원'})

if search:
    result = result[result['예약매장'].str.contains(search)]
if show_loss:
    result = result[result['손실금액'] > 0]

sort_map = {"총건수":"총건수","체결률":"체결률(%)","체결금액":"체결금액","손실금액":"손실금액"}
result = result.sort_values(sort_map[sort_col], ascending=False).reset_index(drop=True)

result_display = result.copy()
result_display['체결금액'] = result_display['체결금액'].apply(lambda x: f"{x:,}")
result_display['손실금액'] = result_display['손실금액'].apply(lambda x: f"{x:,}")
result_display['체결률(%)'] = result_display['체결률(%)'].apply(lambda x: f"{x}%")

st.dataframe(result_display[['예약매장','총건수','체결','불가','체결률(%)','체결 성인','체결 초등','체결 미취학','체결 총인원','체결금액','손실금액']], use_container_width=True, hide_index=True)

st.markdown("<br>", unsafe_allow_html=True)
def to_excel(df_raw, result):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_raw.drop(columns=['인입일_dt'], errors='ignore').to_excel(writer, sheet_name='원본데이터', index=False)
        result.to_excel(writer, sheet_name='매장별현황', index=False)
    return output.getvalue()

st.download_button(label="⬇️ 엑셀로 다운로드", data=to_excel(df, result), file_name="예약현황.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
