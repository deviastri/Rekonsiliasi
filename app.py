import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime

st.set_page_config(page_title="ASDP Dashboard", layout="wide")

MENU_ITEMS = [
    "Dashboard",
    "Tiket Terjual",
    "Penambahan & Pengurangan",
    "Naik/Turun Golongan",
    "Rekonsiliasi",
    "Riwayat Upload"
]
menu = st.sidebar.radio("Pilih Halaman", MENU_ITEMS)

data_dir = "asdp_dashboard_final/data"
os.makedirs(data_dir, exist_ok=True)

def save_excel(df, name):
    df.to_excel(f"{data_dir}/{name}.xlsx", index=False)

def load_excel(name):
    path = f"{data_dir}/{name}.xlsx"
    if os.path.exists(path):
        return pd.read_excel(path)
    return pd.DataFrame()

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
    df_tiket = load_excel("tiket_terjual")
    df_penambahan = load_excel("penambahan")
    df_pengurangan = load_excel("pengurangan")
    df_golongan = load_excel("golongan")

    tiket = df_tiket[
        (df_tiket['Tanggal Mulai'] >= tgl_tiket_start) &
        (df_tiket['Tanggal Selesai'] <= tgl_tiket_end)
    ].groupby('Pelabuhan Asal')['Jumlah'].sum().reindex(pelabuhan, fill_value=0) if not df_tiket.empty else pd.Series(0, index=pelabuhan)

    penambahan = df_penambahan[df_penambahan['Tanggal'] == tgl_penambahan].groupby('Pelabuhan Asal')['Penambahan'].sum().reindex(pelabuhan, fill_value=0) if not df_penambahan.empty else pd.Series(0, index=pelabuhan)

    pengurangan = df_pengurangan[df_pengurangan['Tanggal'] == tgl_pengurangan].groupby('Pelabuhan Asal')['Pengurangan'].sum().reindex(pelabuhan, fill_value=0) if not df_pengurangan.empty else pd.Series(0, index=pelabuhan)

    if not df_golongan.empty:
        df_golongan = df_golongan[df_golongan['Pelabuhan Asal'] != 'TOTAL']
        df_golongan = df_golongan.set_index('Pelabuhan Asal').reindex(pelabuhan, fill_value=0)
        golongan_sum = df_golongan['Selisih Naik/Turun Golongan'].str.replace("Rp ", "").str.replace(".", "").astype(int)
    else:
        golongan_sum = pd.Series(0, index=pelabuhan)

    pinbuk = tiket + penambahan - pengurangan + golongan_sum
    df_rekap = pd.DataFrame({
        'Pelabuhan Asal': pelabuhan,
        'Tiket Terjual': tiket.values,
        'Penambahan': penambahan.values,
        'Pengurangan': pengurangan.values,
        'Naik/Turun Golongan': golongan_sum.values,
        'Nominal Pinbuk': pinbuk.values
    })

    for col in df_rekap.columns[1:]:
        df_rekap[col] = df_rekap[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

    total = {
        'Pelabuhan Asal': 'TOTAL',
        'Tiket Terjual': df_rekap['Tiket Terjual'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Penambahan': df_rekap['Penambahan'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Pengurangan': df_rekap['Pengurangan'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Naik/Turun Golongan': df_rekap['Naik/Turun Golongan'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Nominal Pinbuk': df_rekap['Nominal Pinbuk'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum()
    }

    for key in total:
        if key != 'Pelabuhan Asal':
            total[key] = f"Rp {total[key]:,.0f}".replace(",", ".")

    df_rekap = pd.concat([df_rekap, pd.DataFrame([total])], ignore_index=True)
    st.dataframe(df_rekap, use_container_width=True)

elif menu == "Tiket Terjual":
    st.title("🎟️ Tiket Terjual")
    f = st.file_uploader("Upload File Tiket Terjual (.xlsx)", type="xlsx")
    if f:
        try:
            df_raw = pd.read_excel(f, header=None)
            pelabuhan = str(df_raw.iloc[2, 1]).strip().upper()
            jumlah_data = pd.to_numeric(df_raw.iloc[11:, 4].dropna(), errors='coerce')
            jumlah = jumlah_data.iloc[-1] if not jumlah_data.empty else 0
            tanggal = pd.to_datetime(df_raw.iloc[11:, 2].dropna(), errors='coerce')
            df_save = pd.DataFrame([{
                'Pelabuhan Asal': pelabuhan,
                'Jumlah': int(jumlah),
                'Tanggal Mulai': tanggal.min().date(),
                'Tanggal Selesai': tanggal.max().date()
            }])
            save_excel(df_save, "tiket_terjual")
            st.success("✅ Data berhasil disimpan.")
            st.dataframe(df_save, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")

elif menu == "Penambahan & Pengurangan":
    st.title("📈 Penambahan dan 📉 Pengurangan")
    f = st.file_uploader("Upload File Boarding Pass (.xlsx)", type="xlsx")
    if f:
        try:
            df = pd.read_excel(f)
            df.columns = df.columns.str.upper().str.strip()
            df['JAM'] = pd.to_numeric(df['JAM'], errors='coerce')
            df['TARIF'] = pd.to_numeric(df['TARIF'], errors='coerce')
            df['CETAK BOARDING PASS'] = pd.to_datetime(df['CETAK BOARDING PASS'], errors='coerce')
            df['ASAL'] = df['ASAL'].str.upper().str.strip()
            df = df[df['JAM'].between(0, 7)]
            col1, col2 = st.columns(2)
            with col1:
                tgl_p = st.date_input("Tanggal Penambahan")
            with col2:
                tgl_m = st.date_input("Tanggal Pengurangan")
            pelabuhan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            df_p = df[df['CETAK BOARDING PASS'].dt.date == tgl_p]
            df_m = df[df['CETAK BOARDING PASS'].dt.date == tgl_m]
            p_sum = df_p.groupby('ASAL')['TARIF'].sum().reindex(pelabuhan, fill_value=0)
            m_sum = df_m.groupby('ASAL')['TARIF'].sum().reindex(pelabuhan, fill_value=0)
            df_final = pd.DataFrame({
                'Pelabuhan Asal': pelabuhan,
                'Penambahan': p_sum.values,
                'Pengurangan': m_sum.values,
                'Tanggal': tgl_p
            })
            st.dataframe(df_final, use_container_width=True)
            save_excel(df_final[['Pelabuhan Asal', 'Penambahan', 'Tanggal']], "penambahan")
            save_excel(df_final[['Pelabuhan Asal', 'Pengurangan', 'Tanggal']], "pengurangan")
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")

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

            tanggal_cols = [col for col in df_inv.columns if 'TANGGAL' in col]
            if tanggal_cols:
                df_inv['TANGGAL'] = pd.to_datetime(df_inv[tanggal_cols[0]], errors='coerce')
            else:
                st.warning("❌ Kolom tanggal tidak ditemukan di file Invoice.")
                st.stop()

            df_tik['CETAK'] = pd.to_datetime(df_tik['CETAK'], errors='coerce')
            tgl_min = df_inv['TANGGAL'].min().date()
            tgl_max = df_inv['TANGGAL'].max().date()
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
            save_excel(df_final, "golongan")
            st.dataframe(df_final, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")

elif menu == "Rekonsiliasi":
    st.title("💸 Rekonsiliasi Invoice vs Rekening")

    f_inv = st.file_uploader("Upload File Invoice", type=["xlsx"], key="rinv")
    f_bank = st.file_uploader("Upload File Rekening Koran", type=["xlsx"], key="rbank")

    if f_inv and f_bank:
        try:
            df_inv = pd.read_excel(f_inv)
            df_bank = pd.read_excel(f_bank, skiprows=11)

            df_inv.columns = df_inv.columns.str.lower().str.strip()
            df_bank.columns = df_bank.columns.str.lower().str.strip()

            df_inv = df_inv[['tanggal invoice', 'harga']].dropna()
            df_inv['tanggal invoice'] = pd.to_datetime(df_inv['tanggal invoice'], errors='coerce')
            df_inv['harga'] = pd.to_numeric(df_inv['harga'], errors='coerce')
            df_inv['tanggal'] = df_inv['tanggal invoice'].dt.date

            df_bank = df_bank[['narasi', 'credit transaction']].dropna()
            df_bank['credit transaction'] = pd.to_numeric(df_bank['credit transaction'], errors='coerce')

            records = []
            for _, row in df_bank.iterrows():
                narasi = str(row['narasi'])
                kredit = row['credit transaction']
                tanggal_r = None
                invoice_total = 0

                match = re.search(r'(20\d{6})\s*[-–]?\s*(20\d{6})?', narasi)
                if match:
                    start = pd.to_datetime(match.group(1), format='%Y%m%d', errors='coerce')
                    end = pd.to_datetime(match.group(2), format='%Y%m%d', errors='coerce') if match.group(2) else start
                    if pd.notnull(start) and pd.notnull(end):
                        rng = pd.date_range(start, end)
                        invoice_total = df_inv[df_inv['tanggal'].isin(rng.date)]['harga'].sum()
                        tanggal_r = start.date()

                if tanggal_r:
                    selisih = invoice_total - kredit
                    records.append({
                        'Tanggal': tanggal_r,
                        'Narasi': narasi,
                        'Nominal Kredit': kredit,
                        'Nominal Invoice': invoice_total,
                        'Selisih': selisih
                    })

            df_rekon = pd.DataFrame(records)
            df_rekon[['Nominal Kredit', 'Nominal Invoice', 'Selisih']] = df_rekon[['Nominal Kredit', 'Nominal Invoice', 'Selisih']].fillna(0)

            for col in ['Nominal Kredit', 'Nominal Invoice', 'Selisih']:
                df_rekon[col] = df_rekon[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

            styled = df_rekon.style.set_properties(
                subset=['Nominal Kredit', 'Nominal Invoice', 'Selisih'],
                **{'text-align': 'right'}
            )

            st.dataframe(styled, use_container_width=True)

        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file invoice dan rekening.")

elif menu == "Riwayat Upload":
    st.title("📁 Riwayat Data Upload")

    data_files = {
        "Tiket Terjual": "tiket_terjual.xlsx",
        "Penambahan": "penambahan.xlsx",
        "Pengurangan": "pengurangan.xlsx",
        "Naik/Turun Golongan": "golongan.xlsx",
        "Rekonsiliasi": "rekonsiliasi.xlsx"
    }

    for label, filename in data_files.items():
        path = f"{data_dir}/{filename}"
        if os.path.exists(path):
            df = pd.read_excel(path)
            st.subheader(label)
            st.dataframe(df, use_container_width=True)
        else:
            st.warning(f"Data untuk '{label}' belum ada.")
