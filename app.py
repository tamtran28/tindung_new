from __future__ import annotations
import io
import re
import typing as t
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st
import requests

# ------------------------------
# UI BASICS
# ------------------------------
st.set_page_config(page_title="CRM4/CRM32 Risk Audit (GitHub mode)", layout="wide")
st.title("🔎 CRM4/CRM32 Risk Audit — GitHub links → Streamlit")
st.caption("Đặt file lên GitHub, dán link, chạy toàn bộ pipeline và xuất Excel.")

# ------------------------------
# HELPERS
# ------------------------------
@st.cache_data(show_spinner=False)
def to_raw_github_url(url: str) -> str:
    """Chuyển link `github.com/.../blob/...` sang `raw.githubusercontent.com/...`
    và loại query `?raw=1` nếu có.
    """
    if not isinstance(url, str):
        return url
    u = url.strip()
    if not u:
        return u
    # Nếu đã là raw
    if "raw.githubusercontent.com" in u:
        return u.split("?", 1)[0]
    # Chuyển từ blob → raw
    if "github.com" in u and "/blob/" in u:
        u = u.replace("https://github.com/", "https://raw.githubusercontent.com/")
        u = u.replace("/blob/", "/")
    # Loại query raw
    if "?" in u:
        u = u.split("?", 1)[0]
    return u


@st.cache_data(show_spinner=False)
def fetch_bytes(url: str, token: str | None = None) -> bytes:
    """Tải file bytes từ URL (hỗ trợ GitHub private với token)."""
    headers = {}
    if token:
        headers["Authorization"] = f"token {token.strip()}"
    u = to_raw_github_url(url)
    resp = requests.get(u, headers=headers, timeout=60)
    resp.raise_for_status()
    return resp.content


@st.cache_data(show_spinner=False)
def read_excel_from_url(url: str, token: str | None = None) -> pd.DataFrame:
    """Đọc Excel từ URL (xls/xlsx). Chọn engine theo phần mở rộng."""
    if not url:
        return pd.DataFrame()
    name = url.lower()
    data = fetch_bytes(url, token=token)
    bio = io.BytesIO(data)
    try:
        if name.endswith(".xls") and not name.endswith(".xlsx"):
            df = pd.read_excel(bio, engine="xlrd")
        else:
            df = pd.read_excel(bio, engine="openpyxl")
    finally:
        bio.seek(0)
    # Chuẩn hoá tên cột
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def parse_links(multiline_text: str) -> list[str]:
    """Tách danh sách link theo từng dòng, bỏ trống & comment (#)."""
    links = []
    for line in (multiline_text or "").splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        links.append(s)
    return links


def safe_num_to_str(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    s = s.dropna().astype("int64").astype(str)
    # map giữ index gốc
    return series.index.to_series().map(s).fillna("")


def ensure_columns(df: pd.DataFrame, cols: t.Iterable[str]) -> bool:
    miss = [c for c in cols if c not in df.columns]
    if miss:
        st.warning(f"Thiếu cột: {', '.join(miss)}")
        return False
    return True


def sum_columns(df: pd.DataFrame, colnames: t.List[str]) -> pd.Series:
    present = [c for c in colnames if c in df.columns]
    if not present:
        return pd.Series([0] * len(df), index=df.index)
    return df[present].sum(axis=1, numeric_only=True)


# ------------------------------
# SIDEBAR — GITHUB LINKS & SETTINGS
# ------------------------------
with st.sidebar:
    st.header("📎 Dán link GitHub (raw hoặc blob)")
    gh_token = st.text_input("GitHub token (tuỳ chọn cho repo private)", type="password")

    st.markdown("**1) Bảng mã**")
    link_mdsd = st.text_input("CODE_MDSDV4.xlsx", placeholder="https://github.com/.../CODE_MDSDV4.xlsx")
    link_loaits = st.text_input("CODE_LOAI TSBD.xlsx", placeholder="https://github.com/.../CODE_LOAI TSBD.xlsx")

    st.markdown("**2) CRM4 / CRM32 (nhiều link, mỗi dòng 1 link)**")
    links_crm4_txt = st.text_area("CRM4_Du_no_theo_tai_san_dam_bao_ALL*.xls(x)")
    links_crm32_txt = st.text_area("RPT_CRM_32*.xls(x)")

    st.markdown("**3) Dữ liệu bổ sung (tuỳ chọn)**")
    link_gn1ty = st.text_input("Giai_ngan_tien_mat_1_ty.xls(x)")
    link_muc17 = st.text_input("MUC17.xlsx")
    link_muc55 = st.text_input("Muc55_*.xlsx (Tất toán)")
    link_muc56 = st.text_input("Muc56_*.xlsx (Giải ngân)")
    link_muc57 = st.text_input("Muc57_*.xlsx (Chậm trả)")

    st.divider()
    st.markdown("**Bộ lọc**")
    chi_nhanh = st.text_input("Tên chi nhánh hoặc SOL (ví dụ: HANOI hoặc 001)")
    dia_ban_raw = st.text_input("Tỉnh/TP của đơn vị đang kiểm toán (cách nhau bằng dấu phẩy)")
    ngay_danh_gia = st.date_input("Ngày đánh giá", value=date(2025, 8, 31))

    run_btn = st.button("🚀 Chạy phân tích")


# ------------------------------
# MAIN FLOW
# ------------------------------

def build_pipeline():
    # ----- Read master/mapping tables -----
    if not link_mdsd or not link_loaits:
        st.error("Cần cung cấp link *CODE_MDSDV4.xlsx* và *CODE_LOAI TSBD.xlsx*.")
        return

    with st.spinner("Đang tải bảng mã từ GitHub..."):
        df_muc_dich_file = read_excel_from_url(link_mdsd, token=gh_token)
        df_code_tsbd_file = read_excel_from_url(link_loaits, token=gh_token)

    # ----- Read CRM4/CRM32 files (multi-links) -----
    links_crm4 = parse_links(links_crm4_txt)
    links_crm32 = parse_links(links_crm32_txt)

    if not links_crm4 or not links_crm32:
        st.error("Cần ít nhất 1 link CRM4 và 1 link CRM32.")
        return

    with st.spinner("Đang tải CRM4/CRM32 từ GitHub..."):
        df_crm4_list = [read_excel_from_url(u, token=gh_token) for u in links_crm4]
        df_crm32_list = [read_excel_from_url(u, token=gh_token) for u in links_crm32]
        df_crm4 = pd.concat(df_crm4_list, ignore_index=True) if df_crm4_list else pd.DataFrame()
        df_crm32 = pd.concat(df_crm32_list, ignore_index=True) if df_crm32_list else pd.DataFrame()

    # ----- Basic cleaning as original -----
    if 'CIF_KH_VAY' in df_crm4.columns:
        try:
            df_crm4['CIF_KH_VAY'] = safe_num_to_str(df_crm4['CIF_KH_VAY'])
        except Exception:
            df_crm4['CIF_KH_VAY'] = df_crm4['CIF_KH_VAY'].astype(str)

    if 'CUSTSEQLN' in df_crm32.columns:
        try:
            df_crm32['CUSTSEQLN'] = safe_num_to_str(df_crm32['CUSTSEQLN'])
        except Exception:
            df_crm32['CUSTSEQLN'] = df_crm32['CUSTSEQLN'].astype(str)

    # ----- Filter by branch/SOL -----
    df_crm4_filtered = df_crm4.copy()
    df_crm32_filtered = df_crm32.copy()
    if chi_nhanh.strip():
        key = chi_nhanh.strip().upper()
        if 'BRANCH_VAY' in df_crm4.columns:
            df_crm4_filtered = df_crm4[df_crm4['BRANCH_VAY'].astype(str).str.upper().str.contains(key, na=False)].copy()
        else:
            st.warning("CRM4 thiếu cột 'BRANCH_VAY' — bỏ qua lọc theo chi nhánh.")
        if 'BRCD' in df_crm32.columns:
            df_crm32_filtered = df_crm32[df_crm32['BRCD'].astype(str).str.upper().str.contains(key, na=False)].copy()
        else:
            st.warning("CRM32 thiếu cột 'BRCD' — bỏ qua lọc theo chi nhánh.")

    st.info(f"Số dòng CRM4 sau lọc: **{len(df_crm4_filtered):,}** | CRM32: **{len(df_crm32_filtered):,}**")

    # ------------------------------
    # Map TSBD loại (df_code_tsbd)
    # ------------------------------
    if not ensure_columns(df_code_tsbd_file, ['CODE CAP 2', 'CODE']):
        return
    df_code_tsbd = df_code_tsbd_file[['CODE CAP 2', 'CODE']].copy()
    df_code_tsbd.columns = ['CAP_2', 'LOAI_TS']
    df_tsbd_code = df_code_tsbd[['CAP_2', 'LOAI_TS']].drop_duplicates()

    if 'CAP_2' in df_crm4_filtered.columns:
        df_crm4_filtered = df_crm4_filtered.merge(df_tsbd_code, how='left', on='CAP_2')
        df_crm4_filtered['LOAI_TS'] = df_crm4_filtered.apply(
            lambda row: 'Không TS' if pd.isna(row.get('CAP_2')) or str(row.get('CAP_2')).strip() == '' else row.get('LOAI_TS'),
            axis=1
        )
        df_crm4_filtered['GHI_CHU_TSBD'] = df_crm4_filtered.apply(
            lambda row: 'MỚI' if str(row.get('CAP_2')).strip() != '' and pd.isna(row.get('LOAI_TS')) else '',
            axis=1
        )
    else:
        st.warning("CRM4 thiếu cột 'CAP_2' — không thể map loại TSBD.")
        df_crm4_filtered['LOAI_TS'] = df_crm4_filtered.get('LOAI_TS', 'Không TS')
        df_crm4_filtered['GHI_CHU_TSBD'] = ''

    # ------------------------------
    # Pivot theo loại TS: Dư nợ & Giá trị TS
    # ------------------------------
    for needed in ['CIF_KH_VAY', 'LOAI_TS']:
        if needed not in df_crm4_filtered.columns:
            st.error(f"CRM4 thiếu cột '{needed}' — dừng.")
            return

    if 'DU_NO_PHAN_BO_QUY_DOI' not in df_crm4_filtered.columns:
        df_crm4_filtered['DU_NO_PHAN_BO_QUY_DOI'] = 0.0
    if 'TS_KW_VND' not in df_crm4_filtered.columns:
        df_crm4_filtered['TS_KW_VND'] = 0.0
    if 'LOAI' not in df_crm4_filtered.columns:
        df_crm4_filtered['LOAI'] = ''

    df_vay_4 = df_crm4_filtered.copy()
    df_vay = df_vay_4[~df_vay_4['LOAI'].isin(['Bao lanh', 'LC'])].copy()

    pivot_ts = df_vay.pivot_table(
        index='CIF_KH_VAY',
        columns='LOAI_TS',
        values='TS_KW_VND',
        aggfunc='sum',
        fill_value=0
    ).add_suffix(' (Giá trị TS)').reset_index()

    pivot_no = df_vay.pivot_table(
        index='CIF_KH_VAY',
        columns='LOAI_TS',
        values='DU_NO_PHAN_BO_QUY_DOI',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    pivot_merge = pivot_no.merge(pivot_ts, on='CIF_KH_VAY', how='left')
    cols_no = [c for c in pivot_no.columns if c != 'CIF_KH_VAY']
    cols_ts = [c for c in pivot_merge.columns if c.endswith('(Giá trị TS)')]
    pivot_merge['DƯ NỢ'] = sum_columns(pivot_merge, cols_no)
    pivot_merge['GIÁ TRỊ TS'] = sum_columns(pivot_merge, cols_ts)

    # Info columns
    info_cols = ['CIF_KH_VAY', 'TEN_KH_VAY', 'CUSTTPCD', 'NHOM_NO']
    for c in info_cols:
        if c not in df_crm4_filtered.columns:
            df_crm4_filtered[c] = ''
    df_info = df_crm4_filtered[info_cols].drop_duplicates(subset='CIF_KH_VAY')

    pivot_final = df_info.merge(pivot_merge, on='CIF_KH_VAY', how='left')
    pivot_final = pivot_final.reset_index().rename(columns={'index': 'STT'})
    pivot_final['STT'] = pivot_final['STT'] + 1

    non_ts_non_no = [c for c in pivot_merge.columns if c not in ['CIF_KH_VAY', 'GIÁ TRỊ TS', 'DƯ NỢ'] and '(Giá trị TS)' not in c]
    ts_cols_sorted = sorted([c for c in pivot_merge.columns if c.endswith('(Giá trị TS)')])
    cols_order = ['STT', 'CUSTTPCD', 'CIF_KH_VAY', 'TEN_KH_VAY', 'NHOM_NO'] + sorted(non_ts_non_no) + ts_cols_sorted + ['DƯ NỢ', 'GIÁ TRỊ TS']
    cols_order = [c for c in cols_order if c in pivot_final.columns]
    pivot_final = pivot_final[cols_order]

    # ------------------------------
    # CRM32: Cấp phê duyệt, cơ cấu, mục đích vay
    # ------------------------------
    if 'CAP_PHE_DUYET' in df_crm32_filtered.columns:
        df_crm32_filtered['MA_PHE_DUYET'] = (
            df_crm32_filtered['CAP_PHE_DUYET'].astype(str).str.split('-').str[0].str.strip().str.zfill(2)
        )
    else:
        df_crm32_filtered['MA_PHE_DUYET'] = ''

    ma_cap_c = [f"{i:02d}" for i in range(1, 8)] + [f"{i:02d}" for i in range(28, 32)]
    list_cif_cap_c = df_crm32_filtered[df_crm32_filtered['MA_PHE_DUYET'].isin(ma_cap_c)].get('CUSTSEQLN', pd.Series([], dtype=str)).unique()

    list_co_cau = ['ACOV1', 'ACOV3', 'ATT01', 'ATT02', 'ATT03', 'ATT04',
                   'BCOV1', 'BCOV2', 'BTT01', 'BTT02', 'BTT03',
                   'CCOV2', 'CCOV3', 'CTT03', 'RCOV3', 'RTT03']
    if 'SCHEME_CODE' in df_crm32_filtered.columns:
        cif_co_cau = df_crm32_filtered[df_crm32_filtered['SCHEME_CODE'].isin(list_co_cau)].get('CUSTSEQLN', pd.Series([], dtype=str)).unique()
    else:
        cif_co_cau = []

    # Map mục đích vay
    if ensure_columns(df_muc_dich_file, ['CODE_MDSDV4', 'GROUP']):
        df_muc_dich_vay = df_muc_dich_file[['CODE_MDSDV4', 'GROUP']].copy()
        df_muc_dich_vay.columns = ['MUC_DICH_VAY_CAP_4', 'MUC DICH']
        if 'MUC_DICH_VAY_CAP_4' in df_crm32_filtered.columns:
            df_crm32_filtered = df_crm32_filtered.merge(df_muc_dich_vay, how='left', on='MUC_DICH_VAY_CAP_4')
            df_crm32_filtered['MUC DICH'] = df_crm32_filtered['MUC DICH'].fillna('(blank)')
            df_crm32_filtered['GHI_CHU_TSBD'] = df_crm32_filtered.apply(
                lambda row: 'MỚI' if str(row.get('MUC_DICH_VAY_CAP_4')).strip() != '' and pd.isna(row.get('MUC DICH')) else '',
                axis=1
            )
        else:
            st.warning("CRM32 thiếu cột 'MUC_DICH_VAY_CAP_4' — không map nhóm mục đích vay.")
            df_crm32_filtered['MUC DICH'] = df_crm32_filtered.get('MUC DICH', '(blank)')
            df_crm32_filtered['GHI_CHU_TSBD'] = ''

    # Pivot mục đích vay
    if 'CUSTSEQLN' in df_crm32_filtered.columns and 'MUC DICH' in df_crm32_filtered.columns:
        if 'DU_NO_QUY_DOI' not in df_crm32_filtered.columns:
            df_crm32_filtered['DU_NO_QUY_DOI'] = 0.0
        pivot_mucdich = df_crm32_filtered.pivot_table(
            index='CUSTSEQLN',
            columns='MUC DICH',
            values='DU_NO_QUY_DOI',
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        pivot_mucdich['DƯ NỢ CRM32'] = pivot_mucdich.drop(columns=['CUSTSEQLN']).sum(axis=1, numeric_only=True)
        pivot_final_CRM32 = pivot_mucdich.rename(columns={'CUSTSEQLN': 'CIF_KH_VAY'})
    else:
        pivot_mucdich = pd.DataFrame()
        pivot_final_CRM32 = pd.DataFrame(columns=['CIF_KH_VAY', 'DƯ NỢ CRM32'])

    pivot_full = pivot_final.merge(pivot_final_CRM32, on='CIF_KH_VAY', how='left')
    pivot_full.fillna(0, inplace=True)
    if 'DƯ NỢ' in pivot_full.columns and 'DƯ NỢ CRM32' in pivot_full.columns:
        pivot_full['LECH'] = pivot_full['DƯ NỢ'] - pivot_full['DƯ NỢ CRM32']
    else:
        pivot_full['LECH'] = 0

    # (blank) từ CRM4 không phải Cho vay/Bảo lãnh/LC
    df_crm4_blank = df_crm4_filtered[~df_crm4_filtered['LOAI'].isin(['Cho vay', 'Bao lanh', 'LC'])].copy()
    du_no_bosung = df_crm4_blank.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI': '(blank)'})
    pivot_full = pivot_full.merge(du_no_bosung, on='CIF_KH_VAY', how='left')
    pivot_full['(blank)'] = pivot_full['(blank)'].fillna(0)
    if 'DƯ NỢ CRM32' in pivot_full.columns:
        cols = list(pivot_full.columns)
        if '(blank)' in cols and 'DƯ NỢ CRM32' in cols:
            cols.insert(cols.index('DƯ NỢ CRM32'), cols.pop(cols.index('(blank)')))
            pivot_full = pivot_full[cols]
        pivot_full['DƯ NỢ CRM32'] = pivot_full['DƯ NỢ CRM32'] + pivot_full['(blank)']
        pivot_full['LECH'] = pivot_full['DƯ NỢ'] - pivot_full['DƯ NỢ CRM32']

    # Cờ nhóm nợ / PD cấp C / Cơ cấu
    pivot_full['Nợ nhóm 2'] = pivot_full.get('NHOM_NO', 0).apply(lambda x: 'x' if str(x).strip() == '2' else '')
    pivot_full['Nợ xấu'] = pivot_full.get('NHOM_NO', 0).apply(lambda x: 'x' if str(x).strip() in ['3', '4', '5'] else '')
    pivot_full['Chuyên gia PD cấp C duyệt'] = pivot_full.get('CIF_KH_VAY', '').apply(lambda x: 'x' if x in list_cif_cap_c else '')
    pivot_full['NỢ CƠ_CẤU'] = pivot_full.get('CIF_KH_VAY', '').apply(lambda x: 'x' if x in cif_co_cau else '')

    # Bảo lãnh & LC
    df_baolanh = df_crm4_filtered[df_crm4_filtered['LOAI'] == 'Bao lanh']
    df_lc = df_crm4_filtered[df_crm4_filtered['LOAI'] == 'LC']
    df_baolanh_sum = df_baolanh.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI': 'DƯ_NỢ_BẢO_LÃNH'})
    df_lc_sum = df_lc.groupby('CIF_KH_VAY', as_index=False)['DU_NO_PHAN_BO_QUY_DOI'].sum().rename(columns={'DU_NO_PHAN_BO_QUY_DOI': 'DƯ_NỢ_LC'})
    if 'DƯ_NỢ_BẢO_LÃNH' in pivot_full.columns:
        pivot_full = pivot_full.drop(columns=['DƯ_NỢ_BẢO_LÃNH'])
    pivot_full = pivot_full.merge(df_baolanh_sum, on='CIF_KH_VAY', how='left')
    if 'DƯ_NỢ_LC' in pivot_full.columns:
        pivot_full = pivot_full.drop(columns=['DƯ_NỢ_LC'])
    pivot_full = pivot_full.merge(df_lc_sum, on='CIF_KH_VAY', how='left')
    pivot_full['DƯ_NỢ_BẢO_LÃNH'] = pivot_full['DƯ_NỢ_BẢO_LÃNH'].fillna(0)
    pivot_full['DƯ_NỢ_LC'] = pivot_full['DƯ_NỢ_LC'].fillna(0)

    # Giải ngân tiền mặt 1 tỷ
    if link_gn1ty:
        try:
            df_giai_ngan = read_excel_from_url(link_gn1ty, token=gh_token)
            if 'KHE_UOC' in df_crm32_filtered.columns:
                df_crm32_filtered['KHE_UOC'] = df_crm32_filtered['KHE_UOC'].astype(str).str.strip()
            if 'CUSTSEQLN' in df_crm32_filtered.columns:
                df_crm32_filtered['CUSTSEQLN'] = df_crm32_filtered['CUSTSEQLN'].astype(str).str.strip()
            if 'FORACID' in df_giai_ngan.columns:
                df_giai_ngan['FORACID'] = df_giai_ngan['FORACID'].astype(str).str.strip()
                df_match = df_crm32_filtered[df_crm32_filtered.get('KHE_UOC', '').isin(df_giai_ngan['FORACID'])].copy()
                ds_cif_tien_mat = df_match.get('CUSTSEQLN', pd.Series([], dtype=str)).unique()
                pivot_full['GIẢI_NGÂN_TIEN_MAT'] = pivot_full['CIF_KH_VAY'].astype(str).isin(pd.Series(ds_cif_tien_mat).astype(str)).map({True: 'x', False: ''})
            else:
                st.warning("File giải ngân 1 tỷ thiếu cột FORACID — bỏ qua cờ GIẢI_NGÂN_TIEN_MAT.")
        except Exception as e:
            st.warning(f"Không đọc được file giải ngân 1 tỷ: {e}")
    else:
        pivot_full['GIẢI_NGÂN_TIEN_MAT'] = pivot_full.get('GIẢI_NGÂN_TIEN_MAT', '')

    # Cầm cố tại TCTD khác (CAP_2 chứa 'TCTD')
    if 'CAP_2' in df_crm4_filtered.columns:
        df_cc_tctd = df_crm4_filtered[df_crm4_filtered['CAP_2'].astype(str).str.contains('TCTD', case=False, na=False)]
        df_cc_flag = df_cc_tctd[['CIF_KH_VAY']].drop_duplicates()
        df_cc_flag['Cầm cố tại TCTD khác'] = 'x'
        pivot_full = pivot_full.merge(df_cc_flag, on='CIF_KH_VAY', how='left')
        pivot_full['Cầm cố tại TCTD khác'] = pivot_full['Cầm cố tại TCTD khác'].fillna('')
    else:
        pivot_full['Cầm cố tại TCTD khác'] = ''

    # Top 10 KHCN/KHDN theo DƯ NỢ
    top10_khcn = pivot_full[pivot_full.get('CUSTTPCD', '') == 'Ca nhan'].nlargest(10, 'DƯ NỢ')['CIF_KH_VAY'] if 'DƯ NỢ' in pivot_full.columns else []
    top10_khdn = pivot_full[pivot_full.get('CUSTTPCD', '') == 'Doanh nghiep'].nlargest(10, 'DƯ NỢ')['CIF_KH_VAY'] if 'DƯ NỢ' in pivot_full.columns else []
    pivot_full['Top 10 dư nợ KHCN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in list(top10_khcn) else '')
    pivot_full['Top 10 dư nợ KHDN'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in list(top10_khdn) else '')

    # Quá hạn định giá TSBD (R34)
    ngay_dt = pd.to_datetime(ngay_danh_gia)
    df_crm4_filtered['VALUATION_DATE'] = pd.to_datetime(df_crm4_filtered.get('VALUATION_DATE'), errors='coerce')
    loai_ts_r34 = ['BĐS', 'MMTB', 'PTVT']
    mask_r34 = df_crm4_filtered.get('LOAI_TS', '').isin(loai_ts_r34)
    df_crm4_filtered.loc[mask_r34, 'SO_NGAY_QUA_HAN'] = (
        (ngay_dt - df_crm4_filtered.loc[mask_r34, 'VALUATION_DATE']).dt.days - 365
    )
    df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS'] == 'BĐS', 'SO_THANG_QUA_HAN'] = (
        ((ngay_dt - df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS'] == 'BĐS', 'VALUATION_DATE']).dt.days / 31) - 18
    )
    df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS'].isin(['MMTB', 'PTVT']), 'SO_THANG_QUA_HAN'] = (
        ((ngay_dt - df_crm4_filtered.loc[df_crm4_filtered['LOAI_TS'].isin(['MMTB', 'PTVT']), 'VALUATION_DATE']).dt.days / 31) - 12
    )
    cif_quahan = df_crm4_filtered[df_crm4_filtered.get('SO_NGAY_QUA_HAN', 0) > 30]['CIF_KH_VAY'].dropna().unique()
    pivot_full['KH có TSBĐ quá hạn định giá'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'X' if x in cif_quahan else '')

    # Mục 17 — TS khác địa bàn
    if link_muc17:
        try:
            df_sol = read_excel_from_url(link_muc17, token=gh_token)
            ds_secu = df_crm4_filtered.get('SECU_SRL_NUM', pd.Series([], dtype=str)).dropna().unique()
            if 'C01' in df_sol.columns:
                df_17_filtered = df_sol[df_sol['C01'].isin(ds_secu)]
            else:
                df_17_filtered = pd.DataFrame()
            if not df_17_filtered.empty:
                df_bds = df_17_filtered[df_17_filtered.get('C02', '').astype(str).str.strip().eq('Bat dong san')].copy()
                if 'SECU_SRL_NUM' in df_crm4.columns:
                    df_bds_matched = df_bds[df_bds['C01'].isin(df_crm4['SECU_SRL_NUM'])].copy()
                else:
                    df_bds_matched = df_bds.copy()
                def extract_tinh_thanh(diachi):
                    if pd.isna(diachi):
                        return ''
                    parts = str(diachi).split(','); return parts[-1].strip().lower() if parts else ''
                if 'C19' in df_bds_matched.columns:
                    df_bds_matched['TINH_TP_TSBD'] = df_bds_matched['C19'].apply(extract_tinh_thanh)
                else:
                    df_bds_matched['TINH_TP_TSBD'] = ''
                dia_ban_kt = [t.strip().lower() for t in (dia_ban_raw or '').split(',') if t.strip()]
                df_bds_matched['CANH_BAO_TS_KHAC_DIABAN'] = df_bds_matched['TINH_TP_TSBD'].apply(
                    lambda x: 'x' if x and (x.strip().lower() not in dia_ban_kt) else ''
                )
                ma_ts_canh_bao = df_bds_matched[df_bds_matched['CANH_BAO_TS_KHAC_DIABAN'] == 'x']['C01'].unique() if 'C01' in df_bds_matched.columns else []
                if 'SECU_SRL_NUM' in df_crm4.columns:
                    cif_canh_bao = df_crm4[df_crm4['SECU_SRL_NUM'].isin(ma_ts_canh_bao)].get('CIF_KH_VAY', pd.Series([], dtype=str)).dropna().unique()
                else:
                    cif_canh_bao = []
                pivot_full['KH có TSBĐ khác địa bàn'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in cif_canh_bao else '')
            else:
                df_bds_matched = pd.DataFrame()
                pivot_full['KH có TSBĐ khác địa bàn'] = ''
        except Exception as e:
            st.warning(f"Không đọc được MUC17: {e}")
            df_bds_matched = pd.DataFrame()
            pivot_full['KH có TSBĐ khác địa bàn'] = ''
    else:
        df_bds_matched = pd.DataFrame()
        pivot_full['KH có TSBĐ khác địa bàn'] = ''

    # Tiêu chí 3 — GN & TT cùng ngày
    if link_muc55 and link_muc56:
        try:
            df_55 = read_excel_from_url(link_muc55, token=gh_token)
            df_56 = read_excel_from_url(link_muc56, token=gh_token)
            cols_tt = ['CUSTSEQLN', 'NMLOC', 'KHE_UOC', 'SOTIENGIAINGAN', 'NGAYGN', 'NGAYDH', 'NGAY_TT', 'LOAITIEN']
            if ensure_columns(df_55, cols_tt):
                df_tt = df_55[cols_tt].copy()
                df_tt.columns = ['CIF', 'TEN_KHACH_HANG', 'KHE_UOC', 'SO_TIEN_GIAI_NGAN_VND', 'NGAY_GIAI_NGAN', 'NGAY_DAO_HAN', 'NGAY_TT', 'LOAI_TIEN_HD']
                df_tt['GIAI_NGAN_TT'] = 'Tất toán'
                df_tt['NGAY'] = pd.to_datetime(df_tt['NGAY_TT'], errors='coerce')
            else:
                df_tt = pd.DataFrame(columns=['CIF', 'NGAY', 'GIAI_NGAN_TT'])
            cols_gn = ['CIF', 'TEN_KHACH_HANG', 'KHE_UOC', 'SO_TIEN_GIAI_NGAN_VND', 'NGAY_GIAI_NGAN', 'NGAY_DAO_HAN', 'LOAI_TIEN_HD']
            if ensure_columns(df_56, cols_gn):
                df_gn = df_56[cols_gn].copy()
                df_gn['GIAI_NGAN_TT'] = 'Giải ngân'
                df_gn['NGAY_GIAI_NGAN'] = pd.to_datetime(df_gn['NGAY_GIAI_NGAN'], errors='coerce')
                df_gn['NGAY_DAO_HAN'] = pd.to_datetime(df_gn['NGAY_DAO_HAN'], errors='coerce')
                df_gn['NGAY'] = df_gn['NGAY_GIAI_NGAN']
            else:
                df_gn = pd.DataFrame(columns=['CIF', 'NGAY', 'GIAI_NGAN_TT'])
            df_gop = pd.concat([df_tt, df_gn], ignore_index=True)
            df_gop = df_gop[df_gop['NGAY'].notna()].sort_values(by=['CIF', 'NGAY', 'GIAI_NGAN_TT'])
            if not df_gop.empty:
                df_count = df_gop.groupby(['CIF', 'NGAY', 'GIAI_NGAN_TT']).size().unstack(fill_value=0).reset_index()
                df_count['CO_CA_GN_VA_TT'] = ((df_count.get('Giải ngân', 0) > 0) & (df_count.get('Tất toán', 0) > 0)).astype(int)
                ds_ca_gn_tt = df_count[df_count['CO_CA_GN_VA_TT'] == 1]['CIF'].astype(str).unique()
                pivot_full['CIF_KH_VAY'] = pivot_full['CIF_KH_VAY'].astype(str)
                pivot_full['KH có cả GNG và TT trong 1 ngày'] = pivot_full['CIF_KH_VAY'].apply(lambda x: 'x' if x in ds_ca_gn_tt else '')
            else:
                df_count = pd.DataFrame(); df_gop = pd.DataFrame()
                pivot_full['KH có cả GNG và TT trong 1 ngày'] = ''
        except Exception as e:
            st.warning(f"Không đọc được Muc55/56: {e}")
            df_count = pd.DataFrame(); df_gop = pd.DataFrame()
            pivot_full['KH có cả GNG và TT trong 1 ngày'] = ''
    else:
        df_count = pd.DataFrame(); df_gop = pd.DataFrame()
        pivot_full['KH có cả GNG và TT trong 1 ngày'] = ''

    # Chậm trả (Mục 57)
    if link_muc57:
        try:
            df_delay = read_excel_from_url(link_muc57, token=gh_token)
            if not df_delay.empty:
                df_delay['NGAY_DEN_HAN_TT'] = pd.to_datetime(df_delay.get('NGAY_DEN_HAN_TT'), errors='coerce')
                df_delay['NGAY_THANH_TOAN'] = pd.to_datetime(df_delay.get('NGAY_THANH_TOAN'), errors='coerce')
                ngay_dt = pd.to_datetime(ngay_danh_gia)
                df_delay['NGAY_THANH_TOAN_FILL'] = df_delay['NGAY_THANH_TOAN'].fillna(ngay_dt)
                df_delay['SO_NGAY_CHAM_TRA'] = (df_delay['NGAY_THANH_TOAN_FILL'] - df_delay['NGAY_DEN_HAN_TT']).dt.days
                mask_period = df_delay['NGAY_DEN_HAN_TT'].dt.year.between(2023, 2025)
                df_delay = df_delay[mask_period].copy()
                tmp = pivot_full.copy().rename(columns={'CIF_KH_VAY': 'CIF_ID'})
                df_delay['CIF_ID'] = df_delay.get('CIF_ID', df_delay.get('CIF', '')).astype(str)
                tmp['CIF_ID'] = tmp['CIF_ID'].astype(str)
                df_delay = df_delay.merge(tmp[['CIF_ID', 'DƯ NỢ', 'NHOM_NO']], on='CIF_ID', how='left')
                df_delay = df_delay[df_delay['NHOM_NO'].astype(str).isin(['1', '1.0'])].copy()
                def cap_cham_tra(days):
                    if pd.isna(days):
                        return None
                    elif days >= 10:
                        return '>=10'
                    elif days >= 4:
                        return '4-9'
                    elif days > 0:
                        return '<4'
                    else:
                        return None
                df_delay['CAP_CHAM_TRA'] = df_delay['SO_NGAY_CHAM_TRA'].apply(cap_cham_tra)
                df_delay = df_delay.dropna(subset=['CAP_CHAM_TRA']).copy()
                df_delay['NGAY'] = pd.to_datetime(df_delay['NGAY_DEN_HAN_TT']).dt.date
                order_map = {'>=10': 0, '4-9': 1, '<4': 2}
                df_delay.sort_values(['CIF_ID', 'NGAY', 'CAP_CHAM_TRA'], key=lambda s: s.map(order_map), inplace=True)
                df_unique = df_delay.drop_duplicates(subset=['CIF_ID', 'NGAY'], keep='first').copy()
                df_dem = df_unique.groupby(['CIF_ID', 'CAP_CHAM_TRA']).size().unstack(fill_value=0)
                df_dem['KH Phát sinh chậm trả > 10 ngày'] = np.where(df_dem.get('>=10', 0) > 0, 'x', '')
                df_dem['KH Phát sinh chậm trả 4-9 ngày'] = np.where((df_dem.get('>=10', 0) == 0) & (df_dem.get('4-9', 0) > 0), 'x', '')
                pivot_full = pivot_full.merge(df_dem[['KH Phát sinh chậm trả > 10 ngày', 'KH Phát sinh chậm trả 4-9 ngày']], left_on='CIF_KH_VAY', right_index=True, how='left')
                pivot_full['KH Phát sinh chậm trả > 10 ngày'] = pivot_full['KH Phát sinh chậm trả > 10 ngày'].fillna('')
                pivot_full['KH Phát sinh chậm trả 4-9 ngày'] = pivot_full['KH Phát sinh chậm trả 4-9 ngày'].fillna('')
            else:
                df_delay = pd.DataFrame()
        except Exception as e:
            st.warning(f"Không đọc được Muc57: {e}")
            df_delay = pd.DataFrame()
    else:
        df_delay = pd.DataFrame()

    # ------------------------------
    # OUTPUT — TABS & DOWNLOAD
    # ------------------------------
    tab1, tab2, tab3, tab4 = st.tabs(["📊 KQ_KH (pivot_full)", "📄 Bảng trung gian", "📦 Tải xuống Excel", "ℹ️ Nhật ký/Schema"])

    with tab1:
        st.subheader("Kết quả tổng hợp theo CIF — KQ_KH")
        st.dataframe(pivot_full, use_container_width=True, height=600)

    with tab2:
        st.markdown("**df_crm4_filtered (LOAI_TS)**")
        st.dataframe(df_crm4_filtered, use_container_width=True, height=300)
        st.markdown("**KQ_CRM4 (pivot_final)**")
        st.dataframe(pivot_final, use_container_width=True, height=300)
        st.markdown("**Pivot_crm4 (pivot_merge)**")
        st.dataframe(pivot_merge, use_container_width=True, height=300)
        st.markdown("**df_crm32_filtered (Mục đích vay)**")
        st.dataframe(df_crm32_filtered, use_container_width=True, height=300)
        st.markdown("**Pivot_crm32 (pivot_mucdich)**")
        st.dataframe(pivot_mucdich, use_container_width=True, height=300)
        if 'df_bds_matched' in locals() and isinstance(df_bds_matched, pd.DataFrame) and not df_bds_matched.empty:
            st.markdown("**Tiêu chí 2_dot3 — TS khác địa bàn (df_bds_matched)**")
            st.dataframe(df_bds_matched, use_container_width=True, height=300)
        if 'df_gop' in locals() and isinstance(df_gop, pd.DataFrame) and not df_gop.empty:
            st.markdown("**Tiêu chí 3_dot3 — Gộp GN/TT (df_gop)**")
            st.dataframe(df_gop, use_container_width=True, height=300)
        if 'df_count' in locals() and isinstance(df_count, pd.DataFrame) and not df_count.empty:
            st.markdown("**Tiêu chí 3_dot3_1 — Đếm theo ngày (df_count)**")
            st.dataframe(df_count, use_container_width=True, height=300)
        if 'df_delay' in locals() and isinstance(df_delay, pd.DataFrame) and not df_delay.empty:
            st.markdown("**Tiêu chí 4 — Chậm trả (df_delay)**")
            st.dataframe(df_delay, use_container_width=True, height=300)

    with tab3:
        st.subheader("Xuất file Excel tổng hợp (nhiều sheet)")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_crm4_filtered.to_excel(writer, sheet_name='df_crm4_LOAI_TS', index=False)
            pivot_final.to_excel(writer, sheet_name='KQ_CRM4', index=False)
            pivot_merge.to_excel(writer, sheet_name='Pivot_crm4', index=False)
            df_crm32_filtered.to_excel(writer, sheet_name='df_crm32_LOAI_TS', index=False)
            pivot_full.to_excel(writer, sheet_name='KQ_KH', index=False)
            if 'pivot_mucdich' in locals() and isinstance(pivot_mucdich, pd.DataFrame) and not pivot_mucdich.empty:
                pivot_mucdich.to_excel(writer, sheet_name='Pivot_crm32', index=False)
            if 'df_delay' in locals() and isinstance(df_delay, pd.DataFrame) and not df_delay.empty:
                df_delay.to_excel(writer, sheet_name='tieu chi 4', index=False)
            if 'df_gop' in locals() and isinstance(df_gop, pd.DataFrame) and not df_gop.empty:
                df_gop.to_excel(writer, sheet_name='tieu chi 3_dot3', index=False)
            if 'df_count' in locals() and isinstance(df_count, pd.DataFrame) and not df_count.empty:
                df_count.to_excel(writer, sheet_name='tieu chi 3_dot3_1', index=False)
            if 'df_bds_matched' in locals() and isinstance(df_bds_matched, pd.DataFrame) and not df_bds_matched.empty:
                df_bds_matched.to_excel(writer, sheet_name='tieu chi 2_dot3', index=False)
        st.download_button(
            label="⬇️ Tải xuống KQ_1710_.xlsx",
            data=buffer.getvalue(),
            file_name="KQ_1710_.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Đã sẵn sàng tải file Excel tổng hợp.")

    with tab4:
        st.markdown(
            """
            **Nhật ký tóm tắt**
            - CRM4 links: `{n4}` file | CRM32 links: `{n32}` file
            - Lọc chi nhánh/SOL: `{sol}`
            - Ngày đánh giá: `{dval}`
            - Tỉnh/TP KT: `{diaban}`

            **Cột quan trọng cần có**
            - CRM4: `CIF_KH_VAY`, `BRANCH_VAY`, `LOAI`, `TS_KW_VND`, `DU_NO_PHAN_BO_QUY_DOI`, `CAP_2`, `TEN_KH_VAY`, `CUSTTPCD`, `NHOM_NO`, `SECU_SRL_NUM`, `VALUATION_DATE`
            - CRM32: `CUSTSEQLN`, `BRCD`, `CAP_PHE_DUYET`, `MUC_DICH_VAY_CAP_4`, `DU_NO_QUY_DOI`, `SCHEME_CODE`, `KHE_UOC`
            - MDSDV4: `CODE_MDSDV4`, `GROUP`
            - LOAI TSBD: `CODE CAP 2`, `CODE`

            *Nếu tên cột chênh lệch, hãy chuẩn hoá trước khi upload lên GitHub hoặc cập nhật đoạn map tương ứng.*
            """.format(
                n4=len(parse_links(links_crm4_txt)),
                n32=len(parse_links(links_crm32_txt)),
                sol=chi_nhanh if chi_nhanh else "(không lọc)",
                dval=ngay_danh_gia,
                diaban=dia_ban_raw or "(trống)",
            )
        )


if run_btn:
    build_pipeline()
else:
    st.info("👈 Dán link GitHub & nhấn **Chạy phân tích**.")
