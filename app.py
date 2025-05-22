import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Dashboard ASDP", layout="wide")

def format_rupiah(x):
    if x < 0:
        return f"- Rp {abs(x):,.0f}".replace(",", ".")
    else:
        return f"Rp {x:,.0f}".replace(",", ".")

def to_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Rekap')
    return buffer.getvalue()

def style_table(df, right_align_cols=[], center_align_cols=[]):
    styler = df.style.set_table_styles([
        {'selector': 'th', 'props': [('font-weight', 'bold'), ('background-color', '#f0f2f6')]},
        {'selector': 'tbody tr:nth-child(odd)', 'props': [('background-color', '#fafafa')]},
    ]).hide_index()
    for col in right_align_cols:
        styler = styler.set_properties(subset=[col], **{'text-align': 'right'})
    for col in center_align_cols:
        styler = styler.set_properties(subset=[col], **{'text-align': 'center'})
    return styler

def menu_dashboard():
    st.title("📊 Dashboard Rekapitulasi Sales Channel")
    st.markdown("Pilih periode laporan untuk tiap komponen:")

    c1, c2 = st.columns(2)
    with c1:
        tgl_tiket_start = st.date_input("Tiket Terjual - Mulai", key="tgl_tt_start")
        tgl_tiket_end = st.date_input("Tiket Terjual - Selesai", key="tgl_tt_end")
    with c2:
        tgl_penambahan = st.date_input("Penambahan - Tanggal", key="tgl_penambahan")
        tgl_pengurangan = st.date_input("Pengurangan - Tanggal", key="tgl_pengurangan")

    c3, c4 = st.columns(2)
    with c3:
        tgl_gol_start = st.date_input("Naik/Turun Golongan - Mulai", key="tgl_gol_start")
    with c4:
        tgl_gol_end = st.date_input("Naik/Turun Golongan - Selesai", key="tgl_gol_end")

    # Dummy data as placeholder, replace with actual data fetch & calculation
    df_rekap = pd.DataFrame({
        'Pelabuhan Asal': ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK'],
        'Tiket Terjual': [10000000, 12000000, 9000000, 11000000],
        'Penambahan': [2000000, 1500000, 500000, 700000],
        'Pengurangan': [1000000, 500000, 300000, 400000],
        'Naik/Turun Golongan': [500000, -200000, 0, 300000],
        'Nominal Pinbuk': [11500000, 13250000, 9200000, 11600000],
    })

    for col in df_rekap.columns[1:]:
        df_rekap[col] = df_rekap[col].apply(format_rupiah)

    st.dataframe(style_table(df_rekap, right_align_cols=df_rekap.columns[1:].tolist()), use_container_width=True)

def menu_tiket_terjual():
    st.title("🎟️ Tiket Terjual")
    st.markdown("Unggah file Excel tiket terjual (boleh banyak).")
    files = st.file_uploader("Upload file Excel", type=["xlsx"], accept_multiple_files=True)

    if files:
        hasil = []
        for f in files:
            try:
                df_raw = pd.read_excel(f, header=None)
                pelabuhan = str(df_raw.iloc[2, 1]).strip().upper()
                jumlah = pd.to_numeric(df_raw.iloc[11:, 4].dropna().iloc[-1], errors='coerce')
                if pd.notnull(jumlah):
                    hasil.append({'Pelabuhan Asal': pelabuhan, 'Jumlah': int(jumlah)})
            except Exception as e:
                st.error(f"Gagal memproses {f.name}: {e}")

        if hasil:
            df = pd.DataFrame(hasil)
            urutan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            df = df.groupby('Pelabuhan Asal')['Jumlah'].sum().reindex(urutan, fill_value=0).reset_index()
            total = df['Jumlah'].sum()
            df['Jumlah'] = df['Jumlah'].apply(format_rupiah)
            total_row = pd.DataFrame([{'Pelabuhan Asal': 'TOTAL', 'Jumlah': format_rupiah(total)}])
            final_df = pd.concat([df, total_row], ignore_index=True)

            st.dataframe(style_table(final_df, right_align_cols=['Jumlah']), use_container_width=True)

            excel_data = to_excel(final_df)
            st.download_button("⬇️ Unduh Excel", excel_data, "tiket_terjual.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Silakan unggah file tiket terjual.")

def menu_penambahan_pengurangan():
    st.title("📈 Penambahan dan 📉 Pengurangan Tarif")
    file = st.file_uploader("Upload File Boarding Pass Excel", type=["xlsx"])
    if file:
        try:
            df = pd.read_excel(file)
            df.columns = df.columns.str.upper().str.strip()

            df['JAM'] = pd.to_numeric(df['JAM'], errors='coerce')
            df['CETAK BOARDING PASS'] = pd.to_datetime(df['CETAK BOARDING PASS'], errors='coerce')
            df['ASAL'] = df['ASAL'].str.upper().str.strip()
            df['TARIF'] = pd.to_numeric(df['TARIF'], errors='coerce')

            tanggal_penambahan = st.date_input("Tanggal Penambahan", key="tgl_penambahan")
            tanggal_pengurangan = st.date_input("Tanggal Pengurangan", key="tgl_pengurangan")

            df_p = df[(df['CETAK BOARDING PASS'].dt.date == tanggal_penambahan) & (df['JAM'].between(0,7))]
            df_m = df[(df['CETAK BOARDING PASS'].dt.date == tanggal_pengurangan) & (df['JAM'].between(0,7))]

            pelabuhan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            p_group = df_p.groupby('ASAL')['TARIF'].sum().reindex(pelabuhan, fill_value=0)
            m_group = df_m.groupby('ASAL')['TARIF'].sum().reindex(pelabuhan, fill_value=0)

            df_final = pd.DataFrame({
                'Pelabuhan Asal': pelabuhan,
                'Penambahan': p_group.values,
                'Pengurangan': m_group.values
            })

            for col in ['Penambahan', 'Pengurangan']:
                df_final[col] = df_final[col].apply(format_rupiah)

            total_row = {
                'Pelabuhan Asal': 'TOTAL',
                'Penambahan': format_rupiah(p_group.sum()),
                'Pengurangan': format_rupiah(m_group.sum())
            }
            df_final = pd.concat([df_final, pd.DataFrame([total_row])], ignore_index=True)

            st.dataframe(style_table(df_final, right_align_cols=['Penambahan', 'Pengurangan']), use_container_width=True)

        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file boarding pass.")

def menu_naik_turun_golongan():
    st.title("🚐 Naik/Turun Golongan")

    uploaded_invoice = st.file_uploader("Upload File Invoice", type=["xlsx"])
    uploaded_ticket = st.file_uploader("Upload File Ticket Summary", type=["xlsx"])

    if uploaded_invoice and uploaded_ticket:
        try:
            df_inv = pd.read_excel(uploaded_invoice, header=1)
            df_tik = pd.read_excel(uploaded_ticket, header=1)

            df_inv.columns = df_inv.columns.str.upper().str.strip()
            df_tik.columns = df_tik.columns.str.upper().str.strip()

            invoice_col = 'NOMER INVOICE' if 'NOMER INVOICE' in df_inv.columns else 'NOMOR INVOICE'
            df_inv['INVOICE'] = df_inv[invoice_col].astype(str).str.strip()
            df_tik['INVOICE'] = df_tik['NOMOR INVOICE'].astype(str).str.strip()

            df_inv['HARGA'] = pd.to_numeric(df_inv['HARGA'], errors='coerce').fillna(0)
            df_tik['TARIF'] = pd.to_numeric(df_tik['TARIF'], errors='coerce').fillna(0) * -1

            df_inv['PELABUHAN'] = df_inv['KEBERANGKATAN'].astype(str).str.upper().str.strip()

            inv_grouped = df_inv.groupby(['INVOICE', 'PELABUHAN'], as_index=False)['HARGA'].sum()
            tik_grouped = df_tik.groupby('INVOICE', as_index=False)['TARIF'].sum()

            merged = pd.merge(inv_grouped, tik_grouped, on='INVOICE', how='left').fillna(0)

            utama = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            merged = merged[merged['PELABUHAN'].isin(utama)]

            merged['SELISIH'] = merged['HARGA'] + merged['TARIF']

            rekap = merged.groupby('PELABUHAN')['SELISIH'].sum().reindex(utama, fill_value=0).reset_index()
            rekap['Keterangan'] = rekap['SELISIH'].apply(lambda x: 'Naik Golongan' if x < 0 else ('Turun Golongan' if x > 0 else ''))
            rekap['Selisih Naik/Turun Golongan'] = rekap['SELISIH'].apply(format_rupiah)

            total = rekap['SELISIH'].sum()
            total_row = pd.DataFrame([{
                'PELABUHAN': 'TOTAL',
                'SELISIH': total,
                'Keterangan': '',
                'Selisih Naik/Turun Golongan': format_rupiah(total)
            }])

            final_df = pd.concat([rekap, total_row], ignore_index=True)
            final_df = final_df[['PELABUHAN', 'Selisih Naik/Turun Golongan', 'Keterangan']]
            final_df.columns = ['Pelabuhan Asal', 'Selisih Naik/Turun Golongan', 'Keterangan']

            st.dataframe(style_table(final_df, right_align_cols=['Selisih Naik/Turun Golongan'], center_align_cols=['Keterangan']), use_container_width=True)

            excel_data = to_excel(final_df)
            st.download_button("⬇️ Unduh Excel", excel_data, "naik_turun_golongan.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file Invoice dan Ticket Summary.")

def menu_rekonsiliasi():
    st.title("💸 Rekonsiliasi Invoice vs Rekening")

    uploaded_invoice = st.file_uploader("Upload File Invoice", type=["xlsx"], key="rinv")
    uploaded_bank = st.file_uploader("Upload File Rekening Koran", type=["xlsx"], key="rbank")

    if uploaded_invoice and uploaded_bank:
        try:
            df_inv = pd.read_excel(uploaded_invoice)
            df_bank = pd.read_excel(uploaded_bank, skiprows=11)

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
                df_rekon[col] = df_rekon[col].apply(format_rupiah)

            styled = df_rekon.style.set_properties(subset=['Nominal Kredit', 'Nominal Invoice', 'Selisih'], **{'text-align': 'right'})
            st.dataframe(styled, use_container_width=True)

        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file invoice dan rekening.")

menu = st.sidebar.selectbox("Menu", ["Dashboard", "Tiket Terjual", "Penambahan & Pengurangan", "Naik/Turun Golongan", "Rekonsiliasi"])

if menu == "Dashboard":
    menu_dashboard()
elif menu == "Tiket Terjual":
    menu_tiket_terjual()
elif menu == "Penambahan & Pengurangan":
    menu_penambahan_pengurangan()
elif menu == "Naik/Turun Golongan":
    menu_naik_turun_golongan()
elif menu == "Rekonsiliasi":
    menu_rekonsiliasi()
