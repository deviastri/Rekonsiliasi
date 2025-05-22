import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="ASDP Dashboard", layout="wide")

def format_rupiah(x):
    return f"Rp {x:,.0f}".replace(",", ".")

MENU_ITEMS = [
    "Dashboard",
    "Tiket Terjual",
    "Penambahan & Pengurangan",
    "Naik/Turun Golongan",
    "Pelimpahan Dana"
]

menu = st.sidebar.radio("Pilih Halaman", MENU_ITEMS)

if menu == "Dashboard":
    st.title("📊 Dashboard Rekonsiliasi Sales Channel - Upload File")

    st.markdown("### Upload File Tiket Terjual (boleh banyak)")
    files_tiket = st.file_uploader("Upload File Tiket Terjual", type=["xlsx"], accept_multiple_files=True)

    st.markdown("### Upload File Boarding Pass (Penambahan & Pengurangan)")
    file_boarding = st.file_uploader("Upload File Boarding Pass (1 file)", type=["xlsx"])

    col1, col2 = st.columns(2)
    with col1:
        tgl_penambahan = st.date_input("Filter Tanggal Penambahan")
    with col2:
        tgl_pengurangan = st.date_input("Filter Tanggal Pengurangan")

    st.markdown("### Upload File Invoice (Naik/Turun Golongan)")
    file_invoice = st.file_uploader("Upload File Invoice", type=["xlsx"])

    st.markdown("### Upload File Ticket Summary (Naik/Turun Golongan)")
    file_ticket_summary = st.file_uploader("Upload File Ticket Summary", type=["xlsx"])

    if files_tiket and file_boarding and file_invoice and file_ticket_summary:
        try:
            # --- Tiket Terjual ---
            hasil = []
            for f in files_tiket:
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
            df_tiket = pd.DataFrame(hasil)
            urutan = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            tiket_sum = df_tiket.groupby('Pelabuhan Asal')['Jumlah'].sum().reindex(urutan, fill_value=0)

            # --- Penambahan & Pengurangan ---
            df_boarding = pd.read_excel(file_boarding)
            df_boarding.columns = df_boarding.columns.str.strip().str.upper()
            df_boarding['JAM'] = pd.to_numeric(df_boarding['JAM'], errors='coerce')
            df_boarding['CETAK BOARDING PASS'] = pd.to_datetime(df_boarding['CETAK BOARDING PASS'], errors='coerce')
            df_boarding['ASAL'] = df_boarding['ASAL'].str.upper().str.strip()
            df_boarding['TARIF'] = pd.to_numeric(df_boarding['TARIF'], errors='coerce')
            df_boarding = df_boarding[df_boarding['JAM'].between(0, 7)]
            df_boarding = df_boarding.dropna(subset=['CETAK BOARDING PASS'])

            df_p = df_boarding[df_boarding['CETAK BOARDING PASS'].dt.date == tgl_penambahan]
            df_m = df_boarding[df_boarding['CETAK BOARDING PASS'].dt.date == tgl_pengurangan]

            p_group = df_p.groupby('ASAL')['TARIF'].sum().reindex(urutan, fill_value=0)
            m_group = df_m.groupby('ASAL')['TARIF'].sum().reindex(urutan, fill_value=0)

            # --- Naik/Turun Golongan ---
            df_inv = pd.read_excel(file_invoice, header=1)
            df_tik = pd.read_excel(file_ticket_summary, header=1)
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
            merged = merged[merged['PELABUHAN'].isin(urutan)]

            merged['SELISIH'] = merged['HARGA'] + merged['TARIF']
            rekap = merged.groupby('PELABUHAN')['SELISIH'].sum().reindex(urutan, fill_value=0).reset_index()

            def format_rp(x):
                return f"- Rp {abs(x):,.0f}".replace(",", ".") if x < 0 else f"Rp {x:,.0f}".replace(",", ".")

            rekap['Keterangan'] = rekap['SELISIH'].apply(lambda x: 'Naik Golongan' if x < 0 else ('Turun Golongan' if x > 0 else ''))
            rekap['Selisih Naik/Turun Golongan'] = rekap['SELISIH'].apply(format_rp)

            # --- Hitung Nominal Pinbuk ---
            nominal_pinbuk = tiket_sum + p_group - m_group + rekap['SELISIH']

            df_final = pd.DataFrame({
                'Pelabuhan Asal': urutan,
                'Tiket Terjual': tiket_sum.values,
                'Penambahan': p_group.values,
                'Pengurangan': m_group.values,
                'Naik/Turun Golongan': rekap['SELISIH'].values,
                'Nominal Pinbuk': nominal_pinbuk.values
            })

            for col in ['Tiket Terjual', 'Penambahan', 'Pengurangan', 'Naik/Turun Golongan', 'Nominal Pinbuk']:
                df_final[col] = df_final[col].apply(lambda x: format_rp(x))

            total_row = {
                'Pelabuhan Asal': 'TOTAL',
                'Tiket Terjual': format_rp(tiket_sum.sum()),
                'Penambahan': format_rp(p_group.sum()),
                'Pengurangan': format_rp(m_group.sum()),
                'Naik/Turun Golongan': format_rp(rekap['SELISIH'].sum()),
                'Nominal Pinbuk': format_rp(nominal_pinbuk.sum())
            }

            df_final = pd.concat([df_final, pd.DataFrame([total_row])], ignore_index=True)
            st.dataframe(df_final, use_container_width=True)

            # Save to session for dashboard persistence
            st.session_state['tiket_terjual'] = df_tiket
            st.session_state['penambahan'] = pd.DataFrame({
                'Pelabuhan Asal': urutan,
                'Penambahan': p_group.values,
                'Tanggal': [tgl_penambahan]*len(urutan)
            })
            st.session_state['pengurangan'] = pd.DataFrame({
                'Pelabuhan Asal': urutan,
                'Pengurangan': m_group.values,
                'Tanggal': [tgl_pengurangan]*len(urutan)
            })
            st.session_state['golongan'] = pd.DataFrame({
                'Pelabuhan Asal': rekap['PELABUHAN'],
                'Selisih Naik/Turun Golongan': rekap['SELISIH'].apply(lambda x: format_rp(x)),
                'Keterangan': rekap['Keterangan']
            })

        except Exception as e:
            st.error(f"Gagal memproses data: {e}")

    else:
        st.info("Silakan upload semua file yang diperlukan.")

# Keep other menus (Tiket Terjual, Penambahan & Pengurangan, Naik/Turun Golongan, Pelimpahan Dana) unchanged
# or implement as needed.


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
    st.title("🚐 Selisih Naik/Turun Golongan")

    uploaded_invoice = st.file_uploader("Upload File Invoice", type=["xlsx"])
    uploaded_ticket = st.file_uploader("Upload File Ticket Summary", type=["xlsx"])

    if uploaded_invoice and uploaded_ticket:
        try:
            # Baca data
            df_inv = pd.read_excel(uploaded_invoice, header=1)
            df_tik = pd.read_excel(uploaded_ticket, header=1)

            # Normalisasi nama kolom
            df_inv.columns = df_inv.columns.str.upper().str.strip()
            df_tik.columns = df_tik.columns.str.upper().str.strip()

            # Kolom nomor invoice
            invoice_col = 'NOMER INVOICE' if 'NOMER INVOICE' in df_inv.columns else 'NOMOR INVOICE'
            df_inv['INVOICE'] = df_inv[invoice_col].astype(str).str.strip()
            df_tik['INVOICE'] = df_tik['NOMOR INVOICE'].astype(str).str.strip()

            # Konversi kolom angka dan kali tarif dengan -1
            df_inv['HARGA'] = pd.to_numeric(df_inv['HARGA'], errors='coerce').fillna(0)
            df_tik['TARIF'] = pd.to_numeric(df_tik['TARIF'], errors='coerce').fillna(0) * -1

            # Pelabuhan dari invoice
            df_inv['PELABUHAN'] = df_inv['KEBERANGKATAN'].astype(str).str.upper().str.strip()

            # Hitung total harga per invoice dan pelabuhan
            inv_grouped = df_inv.groupby(['INVOICE', 'PELABUHAN'], as_index=False)['HARGA'].sum()

            # Hitung total tarif per invoice tiket summary
            tik_grouped = df_tik.groupby('INVOICE', as_index=False)['TARIF'].sum()

            # Gabungkan berdasarkan invoice dan pelabuhan
            merged = pd.merge(inv_grouped, tik_grouped, on='INVOICE', how='left').fillna(0)

            utama = ['MERAK', 'BAKAUHENI', 'KETAPANG', 'GILIMANUK']
            merged = merged[merged['PELABUHAN'].isin(utama)]

            # Hitung selisih
            merged['SELISIH'] = merged['HARGA'] + merged['TARIF']

            # Agregasi selisih per pelabuhan
            rekap = merged.groupby('PELABUHAN')['SELISIH'].sum().reindex(utama, fill_value=0).reset_index()

            # Keterangan naik/turun golongan
            rekap['Keterangan'] = rekap['SELISIH'].apply(
                lambda x: 'Naik Golongan' if x < 0 else ('Turun Golongan' if x > 0 else '')
            )

            # Format Rupiah dengan tanda minus di depan jika negatif
            def format_rp(x):
                if x < 0:
                    return f"- Rp {abs(x):,.0f}".replace(",", ".")
                else:
                    return f"Rp {x:,.0f}".replace(",", ".")

            rekap['Selisih Naik/Turun Golongan'] = rekap['SELISIH'].apply(format_rp)

            # Total keseluruhan
            total = rekap['SELISIH'].sum()
            total_row = pd.DataFrame([{
                'PELABUHAN': 'TOTAL',
                'SELISIH': total,
                'Keterangan': '',
                'Selisih Naik/Turun Golongan': format_rp(total)
            }])

            final_df = pd.concat([rekap, total_row], ignore_index=True)
            final_df = final_df[['PELABUHAN', 'Selisih Naik/Turun Golongan', 'Keterangan']]
            final_df.columns = ['Pelabuhan Asal', 'Selisih Naik/Turun Golongan', 'Keterangan']

            st.dataframe(final_df, use_container_width=True)

        except Exception as e:
            st.error(f"Gagal memproses file: {e}")
    else:
        st.info("Silakan upload file Invoice dan Ticket Summary.")


elif menu == "Pelimpahan Dana":
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
