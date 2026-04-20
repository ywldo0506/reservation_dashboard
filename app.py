import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="단체예약 현황 대시보드", page_icon="📋", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap');
* { font-family: 'Noto Sans KR', sans-serif !important; }
.stApp { background-color: #F0F2F5; }
.dash-header { background:white; padding:18px 28px; border-radius:14px; margin-bottom:20px; box-shadow:0 2px 8px rgba(0,0,0,0.06); display:flex; align-items:center; gap:12px; }
.dash-header h1 { font-size:22px; font-weight:900; color:#1a1a1a; margin:0; }
.dash-badge { background:#2E7D32; color:white; font-size:11px; font-weight:700; padding:4px 12px; border-radius:20px; }
.metric-card { background:white; border-radius:14px; padding:20px 22px; box-shadow:0 2px 8px rgba(0,0,0,0.06); border-top:4px solid #E0E0E0; }
.metric-card.green{border-top-color:#2E7D32} .metric-card.blue{border-top-color:#1565C0} .metric-card.red{border-top-color:#B71C1C} .metric-card.orange{border-top-color:#E65100}
.metric-label{font-size:12px;color:#888;font-weight:500;margin-bottom:6px}
.metric-value{font-size:26px;font-weight:900;color:#1a1a1a;line-height:1}
.metric-value.green{color:#2E7D32} .metric-value.blue{color:#1565C0} .metric-value.red{color:#B71C1C} .metric-value.orange{color:#E65100}
.metric-sub{font-size:12px;color:#aaa;margin-top:4px}
.section-box{background:white;border-radius:14px;padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:20px}
.section-title{font-size:14px;font-weight:700;color:#1a1a1a;margin-bottom:14px}
.bar-row{display:flex;align-items:center;gap:10px;margin-bottom:10px}
.bar-label{font-size:13px;width:55px;text-align:right;color:#888;flex-shrink:0}
.bar-track{flex:1;height:30px;background:#F5F5F5;border-radius:8px;overflow:hidden}
.bar-fill{height:100%;border-radius:8px;display:flex;align-items:center;padding-left:10px;font-size:12px;font-weight:700;color:white}
.bar-outside{font-size:12px;color:#888;flex-shrink:0}
.bar-count{font-size:12px;color:#888;width:35px;text-align:right;flex-shrink:0}
</style>
""", unsafe_allow_html=True)

SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ9DoJn7KHOQI44dFBskzlmZI8Xi1_tixVWJhIM-dZezbVkM8C5WpCEDXrrLS6QEHd28GCy1rjSxNme/pub?gid=1866516420&single=true&output=csv"

@st.cache_data(ttl=300)
def load_data():
    try:
        raw = pd.read_csv(SHEET_URL, skiprows=1, header=0, dtype=str)
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
        if col not in raw.columns: raw[col] = None
    raw = raw.dropna(subset=['인입일'])
    raw = raw[raw['인입일'].astype(str).str.strip() != '']
    def parse_date(s):
        s = str(s).strip()
        for fmt in ['%Y-%m-%d','%m/%d/%Y','%Y/%m/%d','%m-%d-%Y','%d/%m/%Y']:
            try: return pd.to_datetime(s, format=fmt)
            except: pass
        try: return pd.to_datetime(s)
        except: return pd.NaT
    raw['인입일_dt'] = raw['인입일'].apply(parse_date)
    raw = raw[raw['인입일_dt'].notna()]
    raw = raw[raw['예약매장'].notna() & (raw['예약매장'].astype(str).str.strip() != '') & (raw['예약매장'].astype(str).str.strip() != 'nan')]
    raw['예약매장'] = raw['예약매장'].astype(str).str.strip().apply(lambda x: re.sub(r'\s+', ' ', x))
    raw['예약매장'] = raw['예약매장'].replace({'퀸즈 여의도한강공원점':'퀸즈 여의도 한강공원점','퀸즈 구의 이스트폴점':'퀸즈 구의이스트폴점','퀸즈 천안 펜타포트점':'퀸즈 천안펜타포트점'})
    for col in ['성인','초등','미취학']:
        raw[col] = pd.to_numeric(raw[col], errors='coerce').fillna(0).astype(int)
    raw['총인원'] = raw['성인'] + raw['초등'] + raw['미취학']
    raw['구분'] = raw['구분'].replace({'런지':'런치','평일':'런치','공휴일':'주말','무응':'기타','미정':'기타'})
    raw['구분'] = raw['구분'].fillna('기타')
    raw['체결구분'] = raw['체결여부'].apply(lambda x: '체결' if pd.isna(x) or str(x).strip() in ['','보류','nan'] else '불가')
    raw['예상금액'] = raw['성인']*25900 + raw['초등']*12900 + raw['미취학']*7900
    return raw

st.markdown('<div class="dash-header"><span class="dash-badge">ASHLEY QUEENS</span><h1>단체예약 현황 대시보드</h1></div>', unsafe_allow_html=True)

col_ref, _ = st.columns([1, 7])
with col_ref:
    if st.button("🔄 새로고침"):
        st.cache_data.clear()
        st.rerun()

df_all = load_data()
if df_all.empty:
    st.warning("데이터를 불러오지 못했어요.")
    st.stop()

min_date = df_all['인입일_dt'].min().date()
max_date = df_all['인입일_dt'].max().date()
st.markdown('<div class="section-box">', unsafe_allow_html=True)
col_f1, col_f2, _ = st.columns([2, 2, 4])
with col_f1:
    start_date = st.date_input("📅 시작일", value=min_date, min_value=min_date, max_value=max_date)
with col_f2:
    end_date = st.date_input("📅 종료일", value=max_date, min_value=min_date, max_value=max_date)
st.markdown('</div>', unsafe_allow_html=True)

df = df_all[(df_all['인입일_dt'].dt.date >= start_date) & (df_all['인입일_dt'].dt.date <= end_date)]
st.caption(f"총 {len(df)}건 기준 · 구글 시트 실시간 연동 (5분 자동갱신) · {start_date} ~ {end_date}")

체결df = df[df['체결구분'] == '체결']
불가df = df[df['체결구분'] == '불가']
total = len(df)
체결수 = len(체결df)
불가수 = len(불가df)
체결률 = round(체결수/total*100, 1) if total > 0 else 0
체결금액 = int(체결df['예상금액'].sum())
손실금액 = int(불가df['예상금액'].sum())

c1,c2,c3,c4,c5 = st.columns(5)
for col, color, label, value, sub in [
    (c1,"blue","전체 예약",f"{total:,}건",""),
    (c2,"green","체결",f"{체결수:,}건",f"체결률 {체결률}%"),
    (c3,"red","불가",f"{불가수:,}건",""),
    (c4,"orange","체결 예상금액",f"{체결금액:,.0f}원",""),
    (c5,"red","손실 금액",f"{손실금액:,.0f}원",""),
]:
    with col:
        st.markdown(f'<div class="metric-card {color}"><div class="metric-label">{label}</div><div class="metric-value {color}">{value}</div>{"<div class=metric-sub>"+sub+"</div>" if sub else ""}</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📊 구분별 예약 현황</div>', unsafe_allow_html=True)
구분p = df.groupby('구분').agg(건수=('예약매장','count'), 체결=('체결구분', lambda x:(x=='체결').sum())).reset_index()
order = ['런치','주말','디너','기타']
구분p['ord'] = 구분p['구분'].apply(lambda x: order.index(x) if x in order else 99)
구분p = 구분p.sort_values('ord').drop(columns='ord')
colors = {'런치':'#1565C0','주말':'#2E7D32','디너':'#E65100','기타':'#9E9E9E'}
max_count = 구분p['건수'].max() if len(구분p) > 0 else 1
bars_html = ""
for _, row in 구분p.iterrows():
    pct = row['건수'] / max_count * 100
    color = colors.get(row['구분'], '#666')
    show_inside = pct > 22
    inner = f"{int(row['체결'])}체결" if show_inside else ""
    outside = f'<span class="bar-outside">{int(row["체결"])}체결</span>' if not show_inside else ""
    bars_html += f'<div class="bar-row"><div class="bar-label">{row["구분"]}</div><div class="bar-track"><div class="bar-fill" style="width:{pct}%;background:{color}">{inner}</div></div>{outside}<div class="bar-count">{int(row["건수"])}건</div></div>'
st.markdown(bars_html, unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-title">① 매장별 예약 현황 (체결 기준 인원·금액)</div>', unsafe_allow_html=True)
col_s, col_f1, col_f2 = st.columns([3,1,1])
with col_s: search = st.text_input("검색", placeholder="🔍 매장명 검색...", label_visibility="collapsed")
with col_f1: show_loss = st.checkbox("손실 있는 매장만")
with col_f2: sort_col = st.selectbox("정렬", ["총건수","체결률","체결금액","손실금액"], label_visibility="collapsed")

p = df.groupby('예약매장').agg(총건수=('예약매장','count'), 체결=('체결구분', lambda x:(x=='체결').sum()), 불가=('체결구분', lambda x:(x=='불가').sum())).reset_index()
cp = 체결df.groupby('예약매장').agg(성인=('성인','sum'), 초등=('초등','sum'), 미취학=('미취학','sum'), 총인원=('총인원','sum'), 체결금액=('예상금액','sum')).reset_index()
lp = 불가df.groupby('예약매장').agg(손실금액=('예상금액','sum')).reset_index()
result = p.merge(cp, on='예약매장', how='left').merge(lp, on='예약매장', how='left').fillna(0)
for col in ['성인','초등','미취학','총인원','체결금액','손실금액']:
    result[col] = result[col].astype(int)
result['체결률(%)'] = (result['체결']/result['총건수']*100).round(1)
result = result.rename(columns={'성인':'체결 성인','초등':'체결 초등','미취학':'체결 미취학','총인원':'체결 총인원'})
if search: result = result[result['예약매장'].str.contains(search, na=False)]
if show_loss: result = result[result['손실금액'] > 0]
sort_map = {"총건수":"총건수","체결률":"체결률(%)","체결금액":"체결금액","손실금액":"손실금액"}
result = result.sort_values(sort_map[sort_col], ascending=False).reset_index(drop=True)
rd = result.copy()
rd['체결금액'] = rd['체결금액'].apply(lambda x: f"{x:,}")
rd['손실금액'] = rd['손실금액'].apply(lambda x: f"{x:,}")
rd['체결률(%)'] = rd['체결률(%)'].apply(lambda x: f"{x}%")
st.dataframe(rd[['예약매장','총건수','체결','불가','체결률(%)','체결 성인','체결 초등','체결 미취학','체결 총인원','체결금액','손실금액']], use_container_width=True, hide_index=True)
st.markdown('</div>', unsafe_allow_html=True)

def to_excel(df_raw, result):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_raw.drop(columns=['인입일_dt'], errors='ignore').to_excel(writer, sheet_name='원본데이터', index=False)
        result.to_excel(writer, sheet_name='매장별현황', index=False)
    return output.getvalue()

st.download_button(label="⬇️ 엑셀로 다운로드", data=to_excel(df, result), file_name="예약현황.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
