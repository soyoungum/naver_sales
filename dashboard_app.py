from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


DOWNLOADS_DIR = Path(__file__).resolve().parent / 'downloads'


def extract_series_name(product_name: str) -> str:
    text = str(product_name or '').strip()
    match = re.search(r'시디즈\s+([^\s,]+)', text)
    if match:
        return match.group(1).strip()
    return text


def find_latest_excel_files() -> list[Path]:
    if not DOWNLOADS_DIR.exists():
        return []
    files = sorted(DOWNLOADS_DIR.glob('sales-*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
    return files


def load_sales_dataframe(excel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    if '상품명' not in df.columns:
        raise ValueError('엑셀에 상품명 컬럼이 없습니다.')

    if '실제수량' not in df.columns:
        if '결제상품수량' not in df.columns or '환불수량' not in df.columns:
            raise ValueError('결제상품수량/환불수량 컬럼을 찾을 수 없습니다.')
        df['실제수량'] = pd.to_numeric(df['결제상품수량'], errors='coerce').fillna(0) - pd.to_numeric(df['환불수량'], errors='coerce').fillna(0)

    if '실제금액' not in df.columns:
        if '결제금액' not in df.columns or '환불금액' not in df.columns:
            raise ValueError('결제금액/환불금액 컬럼을 찾을 수 없습니다.')
        df['실제금액'] = pd.to_numeric(df['결제금액'], errors='coerce').fillna(0) - pd.to_numeric(df['환불금액'], errors='coerce').fillna(0)

    df['시리즈명'] = df['상품명'].map(extract_series_name)
    df['실제수량'] = pd.to_numeric(df['실제수량'], errors='coerce').fillna(0)
    df['실제금액'] = pd.to_numeric(df['실제금액'], errors='coerce').fillna(0)

    return df


def build_series_summary(df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df.groupby('시리즈명', dropna=False, as_index=False)
        .agg(
            실제수량=('실제수량', 'sum'),
            실제금액=('실제금액', 'sum'),
        )
    )

    total_qty = summary['실제수량'].sum()
    total_amount = summary['실제금액'].sum()

    summary['수량비중'] = (summary['실제수량'] / total_qty).fillna(0)
    summary['금액비중'] = (summary['실제금액'] / total_amount).fillna(0)

    summary = summary.sort_values('실제금액', ascending=False).reset_index(drop=True)
    return summary


def format_summary_for_table(summary: pd.DataFrame) -> pd.DataFrame:
    view_df = summary.copy()
    view_df.index = view_df.index + 1
    view_df.index.name = '순위'

    view_df['실제수량'] = view_df['실제수량'].map(lambda x: f"{x:,.0f}")
    view_df['실제금액'] = view_df['실제금액'].map(lambda x: f"{x:,.0f}")
    view_df['수량비중'] = view_df['수량비중'].map(lambda x: f"{x:.2%}")
    view_df['금액비중'] = view_df['금액비중'].map(lambda x: f"{x:.2%}")
    return view_df


def draw_dashboard(summary: pd.DataFrame, source_name: str) -> None:
    total_qty = summary['실제수량'].sum()
    total_amount = summary['실제금액'].sum()
    series_count = summary['시리즈명'].nunique()

    st.markdown(
        """
        <style>
            .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
            .main-title {font-size: 2rem; font-weight: 800; margin-bottom: 0.2rem;}
            .sub-title {color: #5b6470; margin-bottom: 1.2rem;}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="main-title">RAWDATA Dashboard</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-title">데이터 파일: {source_name}</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric('총 실제수량', f"{total_qty:,.0f}")
    c2.metric('총 실제금액', f"{total_amount:,.0f}")
    c3.metric('시리즈 수', f"{series_count:,}")

    top_n = st.slider('차트 표시 시리즈 수', min_value=5, max_value=30, value=15, step=1)
    top_summary = summary.head(top_n)

    col_left, col_right = st.columns([1.3, 1])

    with col_left:
        fig_amount = px.bar(
            top_summary,
            x='시리즈명',
            y='실제금액',
            title='시리즈별 실제금액 (상위)',
            color='실제금액',
            color_continuous_scale='Tealgrn',
        )
        fig_amount.update_layout(xaxis_title='', yaxis_title='실제금액', coloraxis_showscale=False)
        st.plotly_chart(fig_amount, use_container_width=True)

    with col_right:
        pie_df = top_summary[['시리즈명', '수량비중']].copy()
        pie_df['비중(%)'] = pie_df['수량비중'] * 100
        fig_qty = px.pie(
            pie_df,
            names='시리즈명',
            values='비중(%)',
            title='시리즈별 수량비중 (상위)',
            hole=0.45,
        )
        fig_qty.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_qty, use_container_width=True)

    st.subheader('시리즈별 실제수량/실제금액')
    st.dataframe(format_summary_for_table(summary), use_container_width=True, height=520)

    csv_bytes = summary.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        label='집계 CSV 다운로드',
        data=csv_bytes,
        file_name='series_summary_dashboard.csv',
        mime='text/csv',
    )


def main() -> None:
    st.set_page_config(
        page_title='RAWDATA Dashboard',
        page_icon='📊',
        layout='wide',
    )

    files = find_latest_excel_files()
    if not files:
        st.error('downloads 폴더에 sales-*.xlsx 파일이 없습니다.')
        return

    st.sidebar.header('데이터 선택')
    default_idx = 0
    selected_name = st.sidebar.selectbox('엑셀 파일', [f.name for f in files], index=default_idx)
    selected_path = next(f for f in files if f.name == selected_name)

    try:
        sales_df = load_sales_dataframe(selected_path)
        summary_df = build_series_summary(sales_df)
    except Exception as exc:
        st.exception(exc)
        return

    draw_dashboard(summary_df, selected_path.name)


if __name__ == '__main__':
    main()
