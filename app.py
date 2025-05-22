import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.set_page_config(page_title="ASDP Dashboard", layout="wide")

MENU_ITEMS = [
    "Dashboard",
    "Tiket Terjual",
    "Penambahan & Pengurangan",
    "Naik/Turun Golongan",
    "Rekonsiliasi"
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

    df_tiket = st.session_state.get("tiket_terjual", pd.DataFrame())
    df_penambahan = st.session_state.get("penambahan", pd.DataFrame())
    df_pengurangan = st.session_state.get("pengurangan", pd.DataFrame())
    df_golongan = st.session_state.get("golongan", pd.DataFrame())

    pelabuhan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']

    if not df_tiket.empty:
        df_tiket = df_tiket[(df_tiket['Tanggal Mulai'] >= tgl_tiket_start) & (df_tiket['Tanggal Selesai'] <= tgl_tiket_end)]
        tiket_sum = df_tiket.groupby('Pelabuhan Asal')['Jumlah'].sum().reindex(pelabuhan, fill_value=0)
    else:
        tiket_sum = pd.Series(0, index=pelabuhan)

    if not df_penambahan.empty:
        df_penambahan = df_penambahan[df_penambahan['Tanggal'] == pd.to_datetime(tgl_penambahan).date()]
        penambahan_sum = df_penambahan.groupby('Pelabuhan Asal')['Penambahan'].sum().reindex(pelabuhan, fill_value=0)
    else:
        penambahan_sum = pd.Series(0, index=pelabuhan)

    if not df_pengurangan.empty:
        df_pengurangan = df_pengurangan[df_pengurangan['Tanggal'] == pd.to_datetime(tgl_pengurangan).date()]
        pengurangan_sum = df_pengurangan.groupby('Pelabuhan Asal')['Pengurangan'].sum().reindex(pelabuhan, fill_value=0)
    else:
        pengurangan_sum = pd.Series(0, index=pelabuhan)

    if not df_golongan.empty:
        df_golongan_filter = df_golongan[df_golongan['Pelabuhan Asal'] != 'TOTAL']
        df_golongan_filter = df_golongan_filter.set_index('Pelabuhan Asal').reindex(pelabuhan, fill_value=0)
        gol_sum = df_golongan_filter['Selisih Naik/Turun Golongan'].str.replace("Rp ", "").str.replace(".", "").astype(int)
    else:
        gol_sum = pd.Series(0, index=pelabuhan)

    pinbuk = tiket_sum + penambahan_sum - pengurangan_sum + gol_sum

    df_final = pd.DataFrame({
        'Pelabuhan Asal': pelabuhan,
        'Tiket Terjual': tiket_sum.values,
        'Penambahan': penambahan_sum.values,
        'Pengurangan': pengurangan_sum.values,
        'Naik/Turun Golongan': gol_sum.values,
        'Nominal Pinbuk': pinbuk.values
    })

    for col in df_final.columns[1:]:
        df_final[col] = df_final[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

    total_row = {
        'Pelabuhan Asal': 'TOTAL',
        'Tiket Terjual': df_final['Tiket Terjual'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Penambahan': df_final['Penambahan'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Pengurangan': df_final['Pengurangan'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Naik/Turun Golongan': df_final['Naik/Turun Golongan'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum(),
        'Nominal Pinbuk': df_final['Nominal Pinbuk'].str.replace("Rp ", "").str.replace(".", "").astype(int).sum()
    }

    for key in total_row:
        if key != 'Pelabuhan Asal':
            total_row[key] = f"Rp {total_row[key]:,.0f}".replace(",", ".")

    df_final = pd.concat([df_final, pd.DataFrame([total_row])], ignore_index=True)
    st.dataframe(df_final, use_container_width=True)



elif menu == "Tiket Terjual":
    st.title("🎟️ Tiket Terjual")
    files = st.file_uploader("Upload File Tiket (boleh banyak)", type=["xlsx"], accept_multiple_files=True)
    if files:
        hasil = []
        for f in files:
            try:
                xl = pd.read_excel(f, header=None)
                pelabuhan = str(xl.iloc[2, 1]).strip().upper()
                periode_text = str(xl.iloc[4, 4])
                match = re.search(r"(\d{4}-\d{2}-\d{2})\s*[s.d\-]+\s*(\d{4}-\d{2}-\d{2})", periode_text)
                if match:
                    tgl_mulai = pd.to_datetime(match.group(1)).date()
                    tgl_selesai = pd.to_datetime(match.group(2)).date()
                else:
                    tgl_mulai = tgl_selesai = pd.to_datetime("today").date()
                jumlah_data = pd.to_numeric(xl.iloc[11:, 4].dropna(), errors='coerce')
                jumlah = jumlah_data.iloc[-1] if not jumlah_data.empty else 0
                if pd.notnull(jumlah):
                    hasil.append({
                        'Pelabuhan Asal': pelabuhan,
                        'Jumlah': int(jumlah),
                        'Tanggal Mulai': tgl_mulai,
                        'Tanggal Selesai': tgl_selesai
                    })
            except Exception as e:
                st.error(f"❌ Gagal memproses {f.name}: {e}")
        if hasil:
            df = pd.DataFrame(hasil)
            st.session_state['tiket_terjual'] = df
            urutan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            df_grouped = df.groupby('Pelabuhan Asal')['Jumlah'].sum().reindex(urutan, fill_value=0).reset_index()
            total = df_grouped['Jumlah'].sum()
            df_grouped['Jumlah'] = df_grouped['Jumlah'].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
            df_grouped = pd.concat([df_grouped, pd.DataFrame([{'Pelabuhan Asal': 'TOTAL', 'Jumlah': f"Rp {total:,.0f}".replace(",", ".")}])])
            st.dataframe(df_grouped, use_container_width=True)
        else:
            st.warning("Tidak ada data valid ditemukan.")
    else:
        st.info("Silakan upload file tiket.")

# menu lain akan ditambahkan di langkah berikut




elif menu == "Penambahan & Pengurangan":
    st.title("📈 Penambahan dan 📉 Pengurangan Tarif")
    file = st.file_uploader("Upload File Excel (boarding pass)", type=["xlsx"])
    if file:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.upper()
        df['JAM'] = pd.to_numeric(df['JAM'], errors='coerce')
        df['CETAK BOARDING PASS'] = pd.to_datetime(df['CETAK BOARDING PASS'], errors='coerce')
        df['ASAL'] = df['ASAL'].str.upper().str.strip()
        df['TARIF'] = pd.to_numeric(df['TARIF'], errors='coerce')
        df = df[df['JAM'].between(0, 7)]
        df = df.dropna(subset=['CETAK BOARDING PASS'])

        col1, col2 = st.columns(2)
        with col1:
            tanggal_penambahan = st.date_input("Tanggal Penambahan", key="tgl_penambahan")
        with col2:
            tanggal_pengurangan = st.date_input("Tanggal Pengurangan", key="tgl_pengurangan")

        pelabuhan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
        df_p = df[df['CETAK BOARDING PASS'].dt.date == tanggal_penambahan]
        df_m = df[df['CETAK BOARDING PASS'].dt.date == tanggal_pengurangan]

        p_group = df_p.groupby('ASAL')['TARIF'].sum().reindex(pelabuhan, fill_value=0)
        m_group = df_m.groupby('ASAL')['TARIF'].sum().reindex(pelabuhan, fill_value=0)

        df_final = pd.DataFrame({
            'Pelabuhan Asal': pelabuhan,
            'Penambahan': p_group.values,
            'Pengurangan': m_group.values,
            'Tanggal': tanggal_penambahan
        })
        st.session_state['penambahan'] = df_final[['Pelabuhan Asal', 'Penambahan', 'Tanggal']]
        st.session_state['pengurangan'] = df_final[['Pelabuhan Asal', 'Pengurangan', 'Tanggal']]
        for col in ['Penambahan', 'Pengurangan']:
            df_final[col] = df_final[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

        total_row = {
            'Pelabuhan Asal': 'TOTAL',
            'Penambahan': f"Rp {p_group.sum():,.0f}".replace(",", "."),
            'Pengurangan': f"Rp {m_group.sum():,.0f}".replace(",", "."),
            'Tanggal': ''
        }

        df_final = pd.concat([df_final, pd.DataFrame([total_row])], ignore_index=True)
        st.dataframe(df_final, use_container_width=True)
    else:
        st.info("Silakan upload file boarding pass.")

elif menu == "Naik/Turun Golongan":
    st.title("🚐 Naik/Turun Golongan")
    f_inv = st.file_uploader("Upload File Invoice", type=["xlsx"])
    f_tik = st.file_uploader("Upload File Tiket Summary", type=["xlsx"])

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
            st.session_state['golongan'] = df_final
            st.dataframe(df_final, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file invoice dan tiket.")

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
                match = re.search(r'(20\d{6})\s*[-–]?\s*(20\d{6})?', narasi)
                if match:
                    start = pd.to_datetime(match.group(1), format='%Y%m%d', errors='coerce')
                    end = pd.to_datetime(match.group(2), format='%Y%m%d', errors='coerce') if match.group(2) else start
                    rng = pd.date_range(start, end)
                    invoice_total = df_inv[df_inv['tanggal'].isin(rng.date)]['harga'].sum()
                    records.append({
                        'Tanggal': start.date(),
                        'Narasi': narasi,
                        'Nominal Kredit': kredit,
                        'Nominal Invoice': invoice_total,
                        'Selisih': invoice_total - kredit
                    })
            df_rekon = pd.DataFrame(records)
            for col in ['Nominal Kredit', 'Nominal Invoice', 'Selisih']:
                df_rekon[col] = df_rekon[col].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
            st.dataframe(df_rekon, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file invoice dan rekening.")
