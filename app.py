import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="ASDP Dashboard", layout="wide")

MENU_ITEMS = [
    "Dashboard",
    "Naik/Turun Golongan"
]
menu = st.sidebar.radio("Pilih Halaman", MENU_ITEMS)

if menu == "Dashboard":
    st.title("📊 Dashboard Sales Channel")

    col1, col2 = st.columns(2)
    with col1:
        tgl_tiket_start = st.date_input("Tiket Terjual - Tanggal Mulai")
        tgl_tiket_end = st.date_input("Tiket Terjual - Tanggal Selesai")
    with col2:
        tgl_penambahan = st.date_input("Penambahan - Tanggal")
        tgl_pengurangan = st.date_input("Pengurangan - Tanggal")

    col3, col4 = st.columns(2)
    with col3:
        tgl_gol_start = st.date_input("Naik/Turun Golongan - Tanggal Mulai")
    with col4:
        tgl_gol_end = st.date_input("Naik/Turun Golongan - Tanggal Selesai")

    pelabuhan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']

    def read_excel(path):
        if os.path.exists(path):
            return pd.read_excel(path)
        return pd.DataFrame()

    df_golongan = read_excel("asdp_dashboard_final/data/golongan.xlsx")

    if not df_golongan.empty:
        df_golongan = df_golongan[df_golongan['Pelabuhan Asal'] != 'TOTAL']
        df_golongan = df_golongan.set_index('Pelabuhan Asal').reindex(pelabuhan, fill_value=0)
        gol_sum = df_golongan['Selisih Naik/Turun Golongan'].str.replace("Rp ", "").str.replace(".", "").astype(int)
    else:
        gol_sum = pd.Series(0, index=pelabuhan)

    df_final = pd.DataFrame({
        'Pelabuhan Asal': pelabuhan,
        'Naik/Turun Golongan': gol_sum.values
    })
    df_final['Naik/Turun Golongan'] = df_final['Naik/Turun Golongan'].apply(lambda x: "Rp {:,.0f}".format(x).replace(",", "."))
    st.dataframe(df_final, use_container_width=True)

elif menu == "Naik/Turun Golongan":
    st.title("🚐 Naik/Turun Golongan")

    f_inv = st.file_uploader("Upload File Invoice", type=["xlsx"], key="gol_inv")
    f_tik = st.file_uploader("Upload File Tiket Summary", type=["xlsx"], key="gol_tik")

    if f_inv and f_tik:
        try:
            df_inv = pd.read_excel(f_inv, header=1)
            df_tik = pd.read_excel(f_tik, header=1)

            df_inv.columns = df_inv.columns.str.upper().str.strip()
            df_tik.columns = df_tik.columns.str.upper().str.strip()

            invoice_col = 'NOMER INVOICE' if 'NOMER INVOICE' in df_inv.columns else 'NOMOR INVOICE'
            df_inv['INVOICE'] = df_inv[invoice_col].astype(str).str.strip()
            df_tik['INVOICE'] = df_tik['NOMOR INVOICE'].astype(str).str.strip()

            df_inv['NILAI'] = pd.to_numeric(df_inv['HARGA'], errors='coerce')
            df_tik['NILAI'] = pd.to_numeric(df_tik['TARIF'], errors='coerce') * -1

            df_inv['TANGGAL'] = pd.to_datetime(df_inv['TANGGAL'], errors='coerce')
            df_tik['CETAK'] = pd.to_datetime(df_tik['CETAK'], errors='coerce')

            tgl_min = df_inv['TANGGAL'].min().date()
            tgl_max = df_inv['TANGGAL'].max().date()
            st.success(f"Periode: {tgl_min.strftime('%d %B %Y')} s.d. {tgl_max.strftime('%d %B %Y')}")

            df_inv = df_inv[df_inv['TANGGAL'].dt.date.between(tgl_min, tgl_max)]
            df_tik = df_tik[df_tik['CETAK'].dt.date.between(tgl_min, tgl_max)]

            df1 = df_inv[['INVOICE', 'KEBERANGKATAN', 'NILAI']].rename(columns={'KEBERANGKATAN': 'Pelabuhan'})
            df2 = df_tik[['INVOICE', 'NILAI']]
            df2['Pelabuhan'] = None

            df_all = pd.concat([df1, df2], ignore_index=True)
            df_all['Pelabuhan'] = df_all['Pelabuhan'].fillna(method='ffill')
            df_all['Pelabuhan'] = df_all['Pelabuhan'].str.upper().str.strip()

            df_group = df_all.groupby(['INVOICE', 'Pelabuhan'])['NILAI'].sum().reset_index()
            df_filtered = df_group[df_group['Pelabuhan'].isin(['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK'])]
            df_sum = df_filtered.groupby('Pelabuhan')['NILAI'].sum().reset_index()
            df_sum = df_sum[df_sum['NILAI'] != 0]
            df_sum['Keterangan'] = df_sum['NILAI'].apply(lambda x: "Turun Golongan" if x > 0 else "Naik Golongan")
            df_sum['Selisih Naik/Turun Golongan'] = df_sum['NILAI'].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

            df_sum = df_sum[['Pelabuhan', 'Selisih Naik/Turun Golongan', 'Keterangan']].rename(columns={'Pelabuhan': 'Pelabuhan Asal'})
            total = df_filtered['NILAI'].sum()
            df_total = pd.DataFrame([{
                'Pelabuhan Asal': 'TOTAL',
                'Selisih Naik/Turun Golongan': f"Rp {total:,.0f}".replace(",", "."),
                'Keterangan': ''
            }])

            df_final = pd.concat([df_sum, df_total], ignore_index=True)
            df_final.to_excel("asdp_dashboard_final/data/golongan.xlsx", index=False)
            st.dataframe(df_final, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
