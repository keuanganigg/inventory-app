import streamlit as st

st.markdown("""
<style>
body, .stApp {
    background: #f8fafc !important;
    color: #1f2937 !important;
    font-family: "Inter", sans-serif;
}
section[data-testid="stSidebar"] > div {
    background: #ffffff !important;
    border-right: 1px solid #e5e7eb;
    padding: 18px;
}
.metric-card {
    background: #ffffff;
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.08);
}
.metric-title { color:#6b7280; font-weight:600; }
.metric-value { font-size:26px; font-weight:800; color:#1f2937; }
</style>
""", unsafe_allow_html=True)

import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import time
from io import BytesIO
import os
import sys
import re
import json

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
except Exception:
        # Jika lib belum terinstal saat pengembangan lokal, biarkan — Streamlit Cloud akan menginstall dari requirements.
        pass

import io as _io_temp

def _get_drive_service():
    sa_json = st.secrets.get("GDRIVE_SERVICE_ACCOUNT", None)
    if not sa_json:
        st.error("Service account credentials tidak ditemukan di st.secrets['GDRIVE_SERVICE_ACCOUNT']. Tambahkan di Streamlit Cloud → Manage app → Secrets.")
        st.stop()
    try:
        sa_info = json.loads(sa_json)
    except Exception as e:
        st.error("Invalid JSON di GDRIVE_SERVICE_ACCOUNT: " + str(e))
        st.stop()
    creds = service_account.Credentials.from_service_account_info(sa_info, scopes=["https://www.googleapis.com/auth/drive"])
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    return service

def download_db_from_drive(file_id, local_path):
    try:
        drive = _get_drive_service()
        request = drive.files().get_media(fileId=file_id)
        fh = _io_temp.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        with open(local_path, "wb") as f:
            f.write(fh.getbuffer())
        return True, "Downloaded DB from Drive."
    except Exception as e:
        return False, str(e)

def upload_db_to_drive(file_id, local_path):
    try:
        drive = _get_drive_service()
        media = MediaFileUpload(local_path, mimetype="application/x-sqlite3", resumable=True)
        updated = drive.files().update(fileId=file_id, media_body=media).execute()
        return True, "Uploaded DB to Drive."
    except Exception as e:
        return False, str(e)

def upload_after_write(local_db_path='inventory_rumah.db'):
    DRIVE_FILE_ID = st.secrets.get("DRIVE_FILE_ID", None)
    if not DRIVE_FILE_ID:
        st.warning("DRIVE_FILE_ID tidak ada di secrets; melewatkan upload_after_write.")
        return
    ok, msg = upload_db_to_drive(DRIVE_FILE_ID, local_db_path)
    if not ok:
        st.warning("Auto-upload gagal: " + str(msg))
    else:
        # gunakan st.toast() jika tersedia, atau st.success()
        try:
            st.toast("Database auto-synced to Drive.")
        except Exception:
            st.success("Database auto-synced to Drive.")

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

DB_PATH = get_resource_path('inventory_rumah.db')

st.set_page_config(
    page_title="Inventory Gudang",
    page_icon="📦",
    layout="wide"
)

# ---------- STARTUP: pastikan DB ada lokal dengan mendownload dari Drive ----------
DRIVE_FILE_ID = st.secrets.get("DRIVE_FILE_ID", None)
LOCAL_DB = "inventory_rumah.db"
if DRIVE_FILE_ID and not os.path.exists(LOCAL_DB):
    ok, msg = download_db_from_drive(DRIVE_FILE_ID, LOCAL_DB)
    if ok:
        st.info("Database berhasil didownload dari Google Drive saat startup.")
    else:
        st.warning("Gagal download DB dari Drive saat startup: " + str(msg))


# Session state
if 'last_submission' not in st.session_state:
    st.session_state.last_submission = None
if 'form_submitted' not in st.session_state:
    st.session_state.form_submitted = False
if 'submission_success' not in st.session_state:
    st.session_state.submission_success = False
if 'import_config' not in st.session_state:
    st.session_state.import_config = {}
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = {}
if 'import_barang_config' not in st.session_state:
    st.session_state.import_barang_config = {}
if 'selected_sheets_barang' not in st.session_state:
    st.session_state.selected_sheets_barang = {}

# ================= LOGIN & ROLE SYSTEM =================
users = {
    "admin": {"password": "admin123", "role": "editor"},
    "viewer1": {"password": "viewer123", "role": "viewer"},
    "viewer2": {"password": "viewer456", "role": "viewer"}
}

if "user_role" not in st.session_state:
    st.session_state.user_role = None

if st.session_state.user_role is None:
    st.title("🔐 Login Aplikasi Inventory")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in users and users[username]["password"] == password:
            st.session_state.user_role = users[username]["role"]
            st.rerun()
        else:
            st.error("❌ Username atau Password salah")
    st.stop()

# ================= FUNGSI HELPER =================

def format_date_only(df, date_columns):
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').apply(lambda x: x.date() if pd.notna(x) else None)
    return df

def create_excel_download(df, filename_prefix, button_label):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
        worksheet = writer.sheets['Data']
        if not df.empty:
            max_row = len(df)
            max_col = len(df.columns) - 1
            worksheet.autofilter(0, 0, max_row, max_col)
    output.seek(0)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{filename_prefix}_{timestamp}.xlsx"
    st.download_button(
        label=button_label,
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def generate_unit_options():
    units = []
    for letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
        for number in range(1, 17):
            units.append(f"{letter}{number}")
    return units

# ================= DATABASE FUNCTIONS =================

def init_db():
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS barang (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nama_barang TEXT NOT NULL,
                stok INTEGER NOT NULL,
                besaran_stok TEXT NOT NULL,
                gudang TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS peminjaman (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barang_id INTEGER,
                nama_barang TEXT NOT NULL,
                jumlah_pinjam INTEGER NOT NULL,
                tanggal_pinjam DATE NOT NULL,
                unit TEXT NOT NULL,
                besaran_stok TEXT NOT NULL,
                gudang TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (barang_id) REFERENCES barang (id)
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS riwayat_stok (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barang_id INTEGER,
                nama_barang TEXT NOT NULL,
                jumlah_tambah INTEGER NOT NULL,
                stok_sebelum INTEGER NOT NULL,
                stok_sesudah INTEGER NOT NULL,
                gudang TEXT NOT NULL,
                tanggal_tambah TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (barang_id) REFERENCES barang (id)
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS hpp (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                unit TEXT NOT NULL,
                tanggal DATE NOT NULL,
                material TEXT NOT NULL,
                harga REAL NOT NULL,
                keterangan TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )''')

    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)

# ================= HPP FUNCTIONS =================

def read_pengeluaran_material(path, sheet_name="Pengeluaran Material", verbose=True):
    df_raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    
    data_rows = []
    skipped_rows = []

    for idx, row in df_raw.iterrows():
        tanggal_raw = row[1]
        material_raw = row[2]
        unit_raw = row[3]
        harga_raw = row[5]
        
        material_str = str(material_raw).lower().strip()
        harga_str = str(harga_raw).strip()
        
        if material_str in ['nan', 'none', ''] or not material_str:
            continue
        
        header_exact = ['material', 'tanggal', 'keterangan', 'no', 'item']
        if material_str in header_exact:
            skipped_rows.append(f"Row {idx}: Header row")
            continue
        
        if re.match(r'^(jumlah|total|subtotal|grand total|catatan|summary)', material_str):
            skipped_rows.append(f"Row {idx}: Summary row - '{material_str[:50]}'")
            continue
        
        if not harga_str or harga_str.lower() in ["nan", "none", ""]:
            continue
        
        harga_clean = re.sub(r'[^\d.-]', '', harga_str)
        
        try:
            harga = float(harga_clean)
            if harga <= 0 or harga > 100_000_000:
                skipped_rows.append(f"Row {idx}: Harga ekstrem - {harga:,.0f}")
                continue
        except (ValueError, TypeError):
            skipped_rows.append(f"Row {idx}: Invalid harga '{harga_raw}'")
            continue

        tanggal = None
        if pd.notna(tanggal_raw) and str(tanggal_raw).strip():
            try:
                tanggal = pd.to_datetime(tanggal_raw, errors='coerce')
            except:
                pass
        
        material = str(material_raw).strip() if pd.notna(material_raw) else ""
        unit = str(unit_raw).strip() if pd.notna(unit_raw) else ""
        
        data_rows.append([tanggal, material, unit, harga])

    df = pd.DataFrame(data_rows, columns=["Tanggal", "Material", "Unit", "Harga"])
    total_harga = df["Harga"].sum()
    total_rupiah = f"Rp {total_harga:,.0f}".replace(",", ".")

    if verbose:
        st.info(f"📊 Total Rows Terbaca: {len(df)} | Total Harga: {total_rupiah}")

    return df, total_harga

def add_hpp_data(unit, tanggal, material, harga, keterangan=""):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()
    # normalize tanggal to YYYY-MM-DD string for sqlite
    if isinstance(tanggal, pd.Timestamp):
        tanggal = tanggal.strftime('%Y-%m-%d')
    elif isinstance(tanggal, datetime):
        tanggal = tanggal.date().strftime('%Y-%m-%d')
    elif isinstance(tanggal, str):
        try:
            parsed = pd.to_datetime(tanggal, errors='coerce')
            if pd.notna(parsed):
                tanggal = parsed.strftime('%Y-%m-%d')
        except:
            pass
    c.execute("""INSERT INTO hpp (unit, tanggal, material, harga, keterangan)
                 VALUES (?, ?, ?, ?, ?)""", (unit, tanggal, material, harga, keterangan))
    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)

def get_hpp_data(unit=None, start_date=None, end_date=None):
    conn = sqlite3.connect('inventory_rumah.db')
    df = pd.read_sql_query("SELECT * FROM hpp", conn)
    conn.close()
    if df.empty:
        return df

    # --- parsing tanggal: coba beberapa format, fallback ke to_datetime ---
    if 'tanggal' in df.columns:
        # jika tipe sudah datetime, biarkan; jika string, coba parse format yang kita gunakan
        if not pd.api.types.is_datetime64_any_dtype(df['tanggal']):
            # first try YYYY-MM-DD then try DD/MM/YYYY then generic parse
            def safe_parse(x):
                if pd.isna(x):
                    return pd.NaT
                s = str(x).strip()
                for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                    try:
                        return pd.to_datetime(s, format=fmt, errors='raise')
                    except Exception:
                        continue
                # last resort: let pandas infer
                return pd.to_datetime(s, errors='coerce')
            df['tanggal'] = df['tanggal'].apply(safe_parse)
        else:
            df['tanggal'] = pd.to_datetime(df['tanggal'], errors='coerce')

    # --- filter unit ---
    if unit and unit not in ["Semua", "Semua Unit"]:
        df = df[df['unit'] == unit]

    # --- filter tanggal (start_date/end_date may be date objects or strings) ---
    if start_date is not None:
        start_dt = pd.to_datetime(start_date, errors='coerce')
        if pd.notna(start_dt):
            df = df[df['tanggal'] >= start_dt]

    if end_date is not None:
        end_dt = pd.to_datetime(end_date, errors='coerce')
        if pd.notna(end_dt):
            df = df[df['tanggal'] <= end_dt]

    # --- untuk konsistensi tampilan/ekspor: format tanggal ke 'DD/MM/YYYY' ---
    if 'tanggal' in df.columns:
        df['tanggal'] = df['tanggal'].dt.strftime('%d/%m/%Y')

    # sort descending by tanggal (opsional)
    if 'tanggal' in df.columns:
        try:
            df = df.sort_values(by='tanggal', ascending=False)
        except Exception:
            pass

    return df


def delete_hpp(hpp_id):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()
    c.execute("DELETE FROM hpp WHERE id = ?", (hpp_id,))
    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)
    return True, "Data HPP berhasil dihapus"

# ================= BARANG FUNCTIONS =================

def add_barang(nama, stok, besaran, gudang, tanggal_dibuat):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()
    c.execute("INSERT INTO barang (nama_barang, stok, besaran_stok, gudang, created_at) VALUES (?, ?, ?, ?, ?)",
              (nama, stok, besaran, gudang, tanggal_dibuat))
    barang_id = c.lastrowid

    if stok > 0:
        c.execute("""INSERT INTO riwayat_stok
                  (barang_id, nama_barang, jumlah_tambah, stok_sebelum, stok_sesudah, gudang, tanggal_tambah)
                  VALUES (?, ?, ?, ?, ?, ?, ?)""",
                  (barang_id, nama, stok, 0, stok, gudang, tanggal_dibuat))

    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)

def kurangi_stok(barang_id, stok_dikurangi, tanggal_transaksi):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()

    c.execute("SELECT nama_barang, stok, gudang FROM barang WHERE id = ?", (barang_id,))
    barang_data = c.fetchone()

    if barang_data:
        nama_barang, stok_sebelum, gudang = barang_data

        if stok_sebelum < stok_dikurangi:
            conn.close()
            return False, f"Stok tidak mencukupi. Stok tersedia: {stok_sebelum}"

        stok_sesudah = stok_sebelum - stok_dikurangi
        c.execute("UPDATE barang SET stok = stok - ? WHERE id = ?", (stok_dikurangi, barang_id))
        c.execute("""INSERT INTO riwayat_stok
                  (barang_id, nama_barang, jumlah_tambah, stok_sebelum, stok_sesudah, gudang, tanggal_tambah)
                  VALUES (?, ?, ?, ?, ?, ?, ?)""",
                  (barang_id, nama_barang, -stok_dikurangi, stok_sebelum, stok_sesudah, gudang, tanggal_transaksi))

        conn.commit()
        conn.close()
        upload_after_write(LOCAL_DB)
        return True, f"Stok berhasil dikurangi {stok_dikurangi}"

    conn.close()
    return False, "Barang tidak ditemukan"

def update_stok(barang_id, stok_tambahan, tanggal_transaksi):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()

    c.execute("SELECT nama_barang, stok, gudang FROM barang WHERE id = ?", (barang_id,))
    barang_data = c.fetchone()

    if barang_data:
        nama_barang, stok_sebelum, gudang = barang_data
        stok_sesudah = stok_sebelum + stok_tambahan

        c.execute("UPDATE barang SET stok = stok + ? WHERE id = ?", (stok_tambahan, barang_id))
        c.execute("""INSERT INTO riwayat_stok
                  (barang_id, nama_barang, jumlah_tambah, stok_sebelum, stok_sesudah, gudang, tanggal_tambah)
                  VALUES (?, ?, ?, ?, ?, ?, ?)""",
                  (barang_id, nama_barang, stok_tambahan, stok_sebelum, stok_sesudah, gudang, tanggal_transaksi))

        conn.commit()
        conn.close()
        upload_after_write(LOCAL_DB)

def get_barang():
    conn = sqlite3.connect('inventory_rumah.db')
    df = pd.read_sql_query("SELECT * FROM barang ORDER BY nama_barang", conn)
    conn.close()
    return df

def get_barang_by_id(barang_id):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()
    c.execute("SELECT * FROM barang WHERE id = ?", (barang_id,))
    result = c.fetchone()
    conn.close()
    return result

def get_riwayat_stok():
    conn = sqlite3.connect('inventory_rumah.db')
    df = pd.read_sql_query("SELECT * FROM riwayat_stok ORDER BY tanggal_tambah DESC", conn)
    conn.close()
    df = format_date_only(df, ['tanggal_tambah'])
    return df

def delete_barang(barang_id):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()

    c.execute("SELECT COUNT(*) FROM peminjaman WHERE barang_id = ?", (barang_id,))
    has_transactions = c.fetchone()[0] > 0

    if has_transactions:
        conn.close()
        return False, "Barang tidak bisa dihapus karena masih ada riwayat penggunaan"

    c.execute("SELECT nama_barang FROM barang WHERE id = ?", (barang_id,))
    nama_barang = c.fetchone()[0]

    c.execute("DELETE FROM barang WHERE id = ?", (barang_id,))
    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)
    return True, f"Barang '{nama_barang}' berhasil dihapus"

def delete_penggunaan(penggunaan_id):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()
    c.execute("DELETE FROM peminjaman WHERE id = ?", (penggunaan_id,))
    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)
    return True, "Riwayat penggunaan berhasil dihapus"

def delete_riwayat_stok(riwayat_id):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()
    c.execute("DELETE FROM riwayat_stok WHERE id = ?", (riwayat_id,))
    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)
    return True, "Riwayat penambahan stok berhasil dihapus"

def add_peminjaman(barang_id, nama_barang, jumlah, tanggal, unit, besaran, gudang):
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()

    try:
        c.execute("SELECT stok FROM barang WHERE id = ?", (barang_id,))
        current_stock = c.fetchone()

        if not current_stock or current_stock[0] < jumlah:
            conn.close()
            return False, f"Stok tidak mencukupi. Stok tersedia: {current_stock[0] if current_stock else 0}"

        c.execute("""INSERT INTO peminjaman
                    (barang_id, nama_barang, jumlah_pinjam, tanggal_pinjam, unit, besaran_stok, gudang)
                    VALUES (?, ?, ?, ?, ?, ?, ?)""",
                  (barang_id, nama_barang, jumlah, tanggal, unit, besaran, gudang))

        c.execute("UPDATE barang SET stok = stok - ? WHERE id = ?", (jumlah, barang_id))

        conn.commit()
        conn.close()
        upload_after_write(LOCAL_DB)

        return True, f"Berhasil menggunakan {jumlah} {besaran} {nama_barang} untuk unit {unit}"

    except Exception as e:
        conn.rollback()
        conn.close()
        return False, f"Error: {str(e)}"

def get_peminjaman():
    conn = sqlite3.connect('inventory_rumah.db')
    df = pd.read_sql_query("SELECT * FROM peminjaman ORDER BY created_at DESC", conn)
    conn.close()
    df = format_date_only(df, ['tanggal_pinjam', 'created_at'])
    return df

def check_stok_rendah():
    conn = sqlite3.connect('inventory_rumah.db')
    df = pd.read_sql_query("SELECT * FROM barang WHERE stok < 20", conn)
    conn.close()
    return df

def add_sample_data():
    conn = sqlite3.connect('inventory_rumah.db')
    c = conn.cursor()

    c.execute("SELECT COUNT(*) FROM barang")
    if c.fetchone()[0] == 0:
        today = datetime.now().date()
        sample_data = [
            ('Semen', 50, 'Sak', 'Gudang 1', today),
            ('Bata', 15, 'PCS', 'Gudang 1', today),
            ('Paving', 25, 'PCS', 'Gudang 2', today),
            ('Besi', 8, 'PCS', 'Gudang 1', today),
            ('Cat', 30, 'Kaleng', 'Gudang 2', today),
            ('Pasir', 12, 'Sak', 'Gudang 1', today),
        ]

        for item in sample_data:
            add_barang(item[0], item[1], item[2], item[3], item[4])

    conn.commit()
    conn.close()
    upload_after_write(LOCAL_DB)

# Inisialisasi
init_db()
add_sample_data()

# Header aplikasi
st.title("📦 Aplikasi Inventory Gudang")
st.markdown("---")

# Sidebar
st.sidebar.title("📋 Menu Navigasi")

if st.session_state.user_role == "viewer":
    menu = st.sidebar.radio(
        "Pilih Menu:",
        [
            "🏠 Dashboard",
            "📊 Laporan",
            "💰 Laporan HPP",
            "⚠️ Stok Rendah"
        ]
    )
else:
    menu = st.sidebar.radio(
        "Pilih Menu:",
        [
            "🏠 Dashboard",
            "📦 Kelola Barang",
            "📝 Penggunaan",
            "📊 Laporan",
            "💰 Kelola HPP",
            "💰 Laporan HPP",
            "⚠️ Stok Rendah",
            "📥 Import/Export Data"
        ]
    )

st.sidebar.write("---")
if st.sidebar.button("🚪 Logout"):
    st.session_state.user_role = None
    st.rerun()

st.sidebar.write("---")
st.sidebar.caption("Gunakan menu untuk navigasi sistem")

# ================= MENU DASHBOARD =================
if menu == "🏠 Dashboard":
    st.header("🏠 Dashboard Inventory")

    df_barang = get_barang()
    df_peminjaman = get_peminjaman()
    stok_rendah = check_stok_rendah()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        total_item = len(df_barang)
        st.metric("📦 Total Item", total_item)

    with col2:
        total_stok = df_barang['stok'].sum() if not df_barang.empty else 0
        st.metric("📊 Total Stok", total_stok)

    with col3:
        penggunaan_hari_ini = len(df_peminjaman[pd.to_datetime(df_peminjaman['tanggal_pinjam'], errors='coerce').dt.date == datetime.now().date()]) if not df_peminjaman.empty else 0
        st.metric("📝 Penggunaan Hari Ini", penggunaan_hari_ini)

    with col4:
        item_stok_rendah = len(stok_rendah)
        st.metric("⚠️ Stok Rendah", item_stok_rendah, delta_color="inverse")

    if not stok_rendah.empty:
        st.error(f"⚠️ PERINGATAN! Ada {len(stok_rendah)} barang dengan stok kurang dari 20!")
        with st.expander("👁️ Lihat Detail Stok Rendah"):
            st.dataframe(stok_rendah[['nama_barang', 'stok', 'besaran_stok', 'gudang']], use_container_width=True)
    else:
        st.success("✅ Semua stok barang mencukupi!")

    if not df_barang.empty:
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 Distribusi Stok per Gudang")
            stok_gudang = df_barang.groupby('gudang')['stok'].sum().reset_index()
            fig = px.pie(stok_gudang, values='stok', names='gudang',
                         title="Distribusi Total Stok per Gudang",
                         color_discrete_sequence=px.colors.qualitative.Set3)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            st.subheader("📋 Status Stok Semua Barang")
            fig2 = px.bar(df_barang, x='nama_barang', y='stok', color='gudang',
                          title="Jumlah Stok per Barang",
                          labels={'stok': 'Jumlah Stok', 'nama_barang': 'Nama Barang'})
            fig2.add_hline(y=20, line_dash="dash", line_color="red",
                           annotation_text="⚠️ Batas Minimum (20)")
            st.plotly_chart(fig2, use_container_width=True)

    st.subheader("📋 Ringkasan Barang")
    if not df_barang.empty:
        st.dataframe(df_barang[['nama_barang', 'stok', 'besaran_stok', 'gudang']], use_container_width=True)
    else:
        st.info("🔭 Belum ada data barang.")

# ================= MENU KELOLA BARANG =================
elif menu == "📦 Kelola Barang":
    st.header("📦 Kelola Barang")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["➕ Tambah Barang", "👁️ Lihat Barang", "🔄 Tambah Stok", "➖ Kurangi Stok", "🗑️ Hapus Barang", "📜 Riwayat Stok"])

    with tab1:
        st.subheader("➕ Tambah Barang Baru")

        with st.form("form_tambah_barang", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nama_barang = st.text_input("🏷 Nama Barang")
                stok = st.number_input("📊 Stok Awal", min_value=0, value=0, step=1)
                tanggal_dibuat = st.date_input("📅 Tanggal Masuk/Dibuat", value=datetime.now().date())
            with col2:
                besaran_stok = st.text_input("📏 Besaran Stok (contoh: kg, sak, pcs, liter, box)", value="pcs")
                gudang = st.selectbox("🏭 Gudang", ["Gudang 1", "Gudang 2"])

            submitted = st.form_submit_button("➕ Tambah Barang", use_container_width=True)

            if submitted:
                if nama_barang.strip() and besaran_stok.strip():
                    add_barang(nama_barang.strip(), stok, besaran_stok.strip(), gudang, tanggal_dibuat)
                    st.success(f"✅ Barang '{nama_barang}' berhasil ditambahkan!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("❌ Nama barang dan besaran stok harus diisi!")

    with tab2:
        st.subheader("👁️ Daftar Semua Barang")
        display_cols_barang = ['id', 'nama_barang', 'stok', 'besaran_stok', 'gudang']
        df_barang = get_barang()

        if not df_barang.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                filter_gudang = st.selectbox("🏭 Filter Gudang", ["Semua", "Gudang 1", "Gudang 2"])
            with col2:
                search_nama = st.text_input("🔍 Cari Nama Barang")
            with col3:
                show_low_stock = st.checkbox("⚠️ Hanya Stok Rendah")

            df_filtered = df_barang.copy()
            if filter_gudang != "Semua":
                df_filtered = df_filtered[df_filtered['gudang'] == filter_gudang]
            if search_nama:
                df_filtered = df_filtered[df_filtered['nama_barang'].str.contains(search_nama, case=False)]
            if show_low_stock:
                df_filtered = df_filtered[df_filtered['stok'] < 20]

            st.info(f"📊 Menampilkan {len(df_filtered)} dari {len(df_barang)} barang")
            st.dataframe(df_filtered[display_cols_barang], use_container_width=True)

            if not df_filtered.empty:
                create_excel_download(df_filtered[display_cols_barang], "data_barang", "📥 Download Excel")
        else:
            st.info("🔭 Belum ada data barang. Silakan tambah barang baru.")

    with tab3:
        st.subheader("🔄 Tambah Stok Barang")
        st.info("💡 Masukkan jumlah stok yang akan DITAMBAHKAN ke stok saat ini")
        df_barang = get_barang()

        if not df_barang.empty:
            barang_options = {f"{row['nama_barang']} ({row['gudang']}) - Sisa: {row['stok']} {row['besaran_stok']}": row['id']
                            for _, row in df_barang.iterrows()}

            with st.form("form_update_stok"):
                col1, col2 = st.columns(2)
                with col1:
                    selected_barang = st.selectbox("📦 Pilih Barang", list(barang_options.keys()))

                barang_id = barang_options[selected_barang]
                current_barang = get_barang_by_id(barang_id)
                stok_sekarang = current_barang[2]
                satuan = current_barang[3]

                with col2:
                    stok_tambahan = st.number_input("📊 Tambah Stok", min_value=0, value=0, step=1,
                                                     help="Masukkan jumlah yang akan ditambahkan ke stok saat ini")

                tanggal_transaksi = st.date_input("📅 Tanggal Penambahan", value=datetime.now().date())

                submitted = st.form_submit_button("🔄 Tambah Stok", use_container_width=True)

                if submitted:
                    if stok_tambahan > 0:
                        update_stok(barang_id, stok_tambahan, tanggal_transaksi)
                        new_stock = stok_sekarang + stok_tambahan
                        st.success(f"✅ Stok berhasil diupdate dari {stok_sekarang} menjadi {new_stock}!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("❌ Jumlah tambah stok harus lebih dari 0!")
        else:
            st.info("🔭 Belum ada barang untuk diupdate.")

    with tab4:
        st.subheader("➖ Kurangi Stok Barang (Koreksi)")
        st.warning("⚠️ Fitur ini untuk koreksi kesalahan input stok, bukan untuk pencatatan penggunaan!")

        df_barang = get_barang()

        if not df_barang.empty:
            barang_options = {f"{row['nama_barang']} ({row['gudang']}) - Sisa: {row['stok']} {row['besaran_stok']}": row['id']
                            for _, row in df_barang.iterrows()}

            with st.form("form_kurangi_stok"):
                col1, col2 = st.columns(2)
                with col1:
                    selected_barang = st.selectbox("📦 Pilih Barang", list(barang_options.keys()), key="kurangi_barang")

                barang_id = barang_options[selected_barang]
                current_barang = get_barang_by_id(barang_id)
                stok_sekarang = current_barang[2]
                satuan = current_barang[3]

                with col2:
                    stok_dikurangi = st.number_input("📉 Kurangi Stok", min_value=0, max_value=stok_sekarang, value=0, step=1,
                                                     help="Masukkan jumlah yang akan dikurangi dari stok saat ini")

                tanggal_transaksi = st.date_input("📅 Tanggal Pengurangan", value=datetime.now().date())

                submitted = st.form_submit_button("➖ Kurangi Stok", use_container_width=True)

                if submitted:
                    if stok_dikurangi > 0:
                        success, message = kurangi_stok(barang_id, stok_dikurangi, tanggal_transaksi)
                        if success:
                            st.success(f"✅ {message}. Stok sekarang: {stok_sekarang - stok_dikurangi}")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(f"❌ {message}")
                    else:
                        st.error("❌ Jumlah pengurangan stok harus lebih dari 0!")
        else:
            st.info("🔭 Belum ada barang untuk diupdate.")

    with tab5:
        st.subheader("🗑️ Hapus Barang")
        st.warning("⚠️ **PERINGATAN:** Barang yang memiliki riwayat penggunaan tidak bisa dihapus!")

        df_barang = get_barang()

        if not df_barang.empty:
            barang_options = {f"{row['nama_barang']} ({row['gudang']}) - Stok: {row['stok']} {row['besaran_stok']}": row['id']
                            for _, row in df_barang.iterrows()}

            with st.form("form_hapus_barang"):
                selected_barang = st.selectbox("🗑️ Pilih Barang yang akan dihapus", list(barang_options.keys()))

                st.markdown("**Konfirmasi penghapusan:**")
                confirm = st.checkbox("✅ Saya yakin ingin menghapus barang ini")

                submitted = st.form_submit_button("🗑️ HAPUS BARANG", type="secondary", use_container_width=True)

                if submitted:
                    if confirm:
                        barang_id = barang_options[selected_barang]
                        success, message = delete_barang(barang_id)

                        if success:
                            st.success(f"✅ {message}")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(f"❌ {message}")
                    else:
                        st.error("❌ Harap centang konfirmasi untuk menghapus barang!")
        else:
            st.info("🔭 Belum ada barang untuk dihapus.")

    with tab6:
        st.subheader("📜 Riwayat Perubahan Stok")

        tab_view, tab_delete = st.tabs(["👁️ Lihat Riwayat", "🗑️ Hapus Riwayat"])

        with tab_view:
            df_riwayat = get_riwayat_stok()

            if not df_riwayat.empty:
                col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
                with col1:
                    start_date = st.date_input("📅 Dari Tanggal", value=datetime.now().date() - timedelta(days=30), key="riwayat_start")
                with col2:
                    end_date = st.date_input("📅 Sampai Tanggal", value=datetime.now().date(), key="riwayat_end")
                with col3:
                    filter_jenis = st.selectbox("Jenis Transaksi", ["Semua", "Tambah", "Kurang"], key="filter_jenis_stok")
                with col4:
                    search_barang = st.text_input("🔍 Cari Barang", key="riwayat_search")

                mask = (pd.to_datetime(df_riwayat['tanggal_tambah'], errors='coerce').dt.date >= start_date) & (pd.to_datetime(df_riwayat['tanggal_tambah'], errors='coerce').dt.date <= end_date)
                df_filtered = df_riwayat.loc[mask]

                if filter_jenis == "Tambah":
                    df_filtered = df_filtered[df_filtered['jumlah_tambah'] > 0]
                elif filter_jenis == "Kurang":
                    df_filtered = df_filtered[df_filtered['jumlah_tambah'] < 0]

                if search_barang:
                    df_filtered = df_filtered[df_filtered['nama_barang'].str.contains(search_barang, case=False)]

                if not df_filtered.empty:
                    st.info(f"📊 Menampilkan {len(df_filtered)} riwayat perubahan stok")

                    display_df = df_filtered.copy()
                    display_df['Jenis'] = display_df['jumlah_tambah'].apply(lambda x: '➕ Tambah' if x > 0 else '➖ Kurang')
                    display_df['Jumlah'] = display_df['jumlah_tambah'].abs()

                    display_df = display_df.rename(columns={
                        'id': 'ID',
                        'nama_barang': 'Nama Barang',
                        'stok_sebelum': 'Stok Sebelum',
                        'stok_sesudah': 'Stok Sesudah',
                        'gudang': 'Gudang',
                        'tanggal_tambah': 'Tanggal'
                    })

                    display_cols = ['ID', 'Jenis', 'Nama Barang', 'Jumlah', 'Stok Sebelum', 'Stok Sesudah', 'Gudang', 'Tanggal']
                    st.dataframe(display_df[display_cols], use_container_width=True)

                    total_penambahan = display_df[display_df['Jenis'] == '➕ Tambah']['Jumlah'].sum()
                    total_pengurangan = display_df[display_df['Jenis'] == '➖ Kurang']['Jumlah'].sum()

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("📊 Total Stok Ditambahkan", total_penambahan)
                    with col2:
                        st.metric("📉 Total Stok Dikurangi", total_pengurangan)

                    create_excel_download(display_df[display_cols], "riwayat_stok", "📥 Download Excel")
                else:
                    st.info("🔭 Tidak ada riwayat perubahan stok dalam rentang tanggal tersebut.")
            else:
                st.info("🔭 Belum ada riwayat perubahan stok.")

        with tab_delete:
            st.warning("⚠️ **PERHATIAN:** Hapus riwayat hanya untuk koreksi kesalahan input. Stok barang TIDAK akan berubah!")

            df_riwayat = get_riwayat_stok()

            if not df_riwayat.empty:
                riwayat_options = {}
                for _, row in df_riwayat.iterrows():
                    jenis = "Tambah" if row['jumlah_tambah'] > 0 else "Kurang"
                    jumlah = abs(row['jumlah_tambah'])
                    label = f"ID-{row['id']}: {jenis} {row['nama_barang']} ({jumlah}) - {row['tanggal_tambah']}"
                    riwayat_options[label] = row['id']

                with st.form("form_hapus_riwayat_stok"):
                    selected_riwayat = st.selectbox("🗑️ Pilih riwayat yang akan dihapus", list(riwayat_options.keys()))

                    st.markdown("**Konfirmasi penghapusan:**")
                    confirm = st.checkbox("✅ Saya yakin ingin menghapus riwayat ini")

                    submitted = st.form_submit_button("🗑️ HAPUS RIWAYAT", type="secondary", use_container_width=True)

                    if submitted:
                        if confirm:
                            riwayat_id = riwayat_options[selected_riwayat]
                            success, message = delete_riwayat_stok(riwayat_id)

                            if success:
                                st.success(f"✅ {message}")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(f"❌ {message}")
                        else:
                            st.error("❌ Harap centang konfirmasi untuk menghapus riwayat!")
            else:
                st.info("🔭 Belum ada riwayat untuk dihapus.")

# ================= MENU PENGGUNAAN =================
elif menu == "📝 Penggunaan":
    st.header("📝 Kelola Penggunaan")

    tab1, tab2 = st.tabs(["📤 Gunakan Barang", "📜 Riwayat Penggunaan"])

    with tab1:
        st.subheader("📤 Gunakan Barang")

        if st.session_state.get('submission_success', False):
            st.success("✅ Penggunaan berhasil diproses!")
            st.balloons()
            st.session_state.submission_success = False

        df_barang = get_barang()
        df_available = df_barang[df_barang['stok'] > 0] if not df_barang.empty else pd.DataFrame()

        if not df_available.empty:
            barang_options = {f"{row['nama_barang']} ({row['gudang']}) - Tersedia: {row['stok']} {row['besaran_stok']}": row['id']
                            for _, row in df_available.iterrows()}

            with st.form("form_penggunaan_fixed", clear_on_submit=False):
                col1, col2 = st.columns(2)
                with col1:
                    selected_barang = st.selectbox("📦 Pilih Barang", list(barang_options.keys()))
                    jumlah_pinjam = st.number_input("📊 Jumlah Gunakan", min_value=1, value=1, step=1)
                    unit = st.text_input("🏠 Unit", placeholder='gudang barat')
                with col2:
                    tanggal_pinjam = st.date_input("📅 Tanggal Gunakan", value=datetime.now().date())
                    st.write("")

                submitted = st.form_submit_button("📤 Konfirmasi Penggunaan", use_container_width=True)

                if submitted and not st.session_state.get('form_submitted', False):
                    st.session_state.form_submitted = True

                    barang_id = barang_options[selected_barang]
                    barang_data = get_barang_by_id(barang_id)

                    if barang_data:
                        success, message = add_peminjaman(
                            barang_id,
                            barang_data[1],
                            jumlah_pinjam,
                            tanggal_pinjam,
                            unit,
                            barang_data[3],
                            barang_data[4]
                        )

                        if success:
                            st.session_state.submission_success = True
                            st.session_state.form_submitted = False
                            st.rerun()
                        else:
                            st.error(f"❌ {message}")
                            st.session_state.form_submitted = False
                    else:
                        st.error("❌ Barang tidak ditemukan!")
                        st.session_state.form_submitted = False

                elif submitted:
                    st.info("⏳ Penggunaan sedang diproses...")

        else:
            st.warning("⚠️ Tidak ada barang yang tersedia untuk digunakan.")

    with tab2:
        st.subheader("📜 Riwayat Penggunaan")

        tab_view, tab_delete = st.tabs(["👁️ Lihat Riwayat", "🗑️ Hapus Riwayat"])

        with tab_view:
            df_peminjaman = get_peminjaman()

            if not df_peminjaman.empty:
                col1, col2, col3 = st.columns(3)
                with col1:
                    start_date = st.date_input("📅 Dari Tanggal", value=datetime.now().date() - timedelta(days=30))
                with col2:
                    end_date = st.date_input("📅 Sampai Tanggal", value=datetime.now().date())
                with col3:
                    search_barang = st.text_input("🔍 Cari Barang")

                mask = (pd.to_datetime(df_peminjaman['tanggal_pinjam'], errors='coerce').dt.date >= start_date) & (pd.to_datetime(df_peminjaman['tanggal_pinjam'], errors='coerce').dt.date <= end_date)
                df_filtered = df_peminjaman.loc[mask]

                if search_barang:
                    df_filtered = df_filtered[df_filtered['nama_barang'].str.contains(search_barang, case=False)]

                if not df_filtered.empty:
                    st.info(f"📊 Menampilkan {len(df_filtered)} transaksi penggunaan")

                    display_df = df_filtered.copy()
                    display_df = display_df.rename(columns={
                        'id': 'ID',
                        'nama_barang': 'Nama Barang',
                        'jumlah_pinjam': 'Jumlah Penggunaan',
                        'tanggal_pinjam': 'Tanggal Penggunaan',
                        'unit': 'Unit',
                        'besaran_stok': 'Satuan',
                        'gudang': 'Gudang'
                    })

                    st.dataframe(display_df[['ID', 'Nama Barang', 'Jumlah Penggunaan', 'Tanggal Penggunaan', 'Unit', 'Satuan', 'Gudang']], use_container_width=True)

                    total_transaksi = len(df_filtered)
                    total_barang = df_filtered['jumlah_pinjam'].sum()

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("📊 Total Transaksi", total_transaksi)
                    with col2:
                        st.metric("📦 Total Barang Digunakan", total_barang)

                    create_excel_download(display_df[['ID', 'Nama Barang', 'Jumlah Penggunaan', 'Tanggal Penggunaan', 'Unit', 'Satuan', 'Gudang']], "riwayat_penggunaan", "📥 Download Excel")
                else:
                    st.info("🔭 Tidak ada penggunaan dalam rentang tanggal tersebut.")
            else:
                st.info("🔭 Belum ada riwayat penggunaan.")

        with tab_delete:
            st.warning("⚠️ **PERHATIAN:** Hapus riwayat hanya untuk koreksi kesalahan input. Stok barang TIDAK akan dikembalikan!")

            df_peminjaman = get_peminjaman()

            if not df_peminjaman.empty:
                penggunaan_options = {f"ID-{row['id']}: {row['nama_barang']} ({row['jumlah_pinjam']} {row['besaran_stok']}) - Unit {row['unit']} - {row['tanggal_pinjam']}": row['id']
                                     for _, row in df_peminjaman.iterrows()}

                with st.form("form_hapus_penggunaan"):
                    selected_penggunaan = st.selectbox("🗑️ Pilih riwayat penggunaan yang akan dihapus", list(penggunaan_options.keys()))

                    st.markdown("**Konfirmasi penghapusan:**")
                    confirm = st.checkbox("✅ Saya yakin ingin menghapus riwayat ini")

                    submitted = st.form_submit_button("🗑️ HAPUS RIWAYAT", type="secondary", use_container_width=True)

                    if submitted:
                        if confirm:
                            penggunaan_id = penggunaan_options[selected_penggunaan]
                            success, message = delete_penggunaan(penggunaan_id)

                            if success:
                                st.success(f"✅ {message}")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(f"❌ {message}")
                        else:
                            st.error("❌ Harap centang konfirmasi untuk menghapus riwayat!")
            else:
                st.info("🔭 Belum ada riwayat untuk dihapus.")

# Penggunaan
elif menu == "📝 Penggunaan":
    st.header("📝 Kelola Penggunaan")

    tab1, tab2 = st.tabs(["📤 Gunakan Barang", "📜 Riwayat Penggunaan"])

    with tab1:
        st.subheader("📤 Gunakan Barang")

        if st.session_state.get('submission_success', False):
            st.success("✅ Penggunaan berhasil diproses!")
            st.balloons()
            st.session_state.submission_success = False

        df_barang = get_barang()
        df_available = df_barang[df_barang['stok'] > 0] if not df_barang.empty else pd.DataFrame()

        if not df_available.empty:
            barang_options = {f"{row['nama_barang']} ({row['gudang']}) - Tersedia: {row['stok']} {row['besaran_stok']}": row['id']
                            for _, row in df_available.iterrows()}

            with st.form("form_penggunaan_fixed", clear_on_submit=False):
                col1, col2 = st.columns(2)
                with col1:
                    selected_barang = st.selectbox("📦 Pilih Barang", list(barang_options.keys()))
                    jumlah_pinjam = st.number_input("📊 Jumlah Gunakan", min_value=1, value=1, step=1)
                    unit_options = generate_unit_options()
                    unit = st.selectbox("🏠 Digunakan untuk Unit", unit_options)
                with col2:
                    tanggal_pinjam = st.date_input("📅 Tanggal Gunakan", value=datetime.now().date())
                    st.write("")

                submitted = st.form_submit_button("📤 Konfirmasi Penggunaan", use_container_width=True)

                if submitted and not st.session_state.get('form_submitted', False):
                    st.session_state.form_submitted = True

                    barang_id = barang_options[selected_barang]
                    barang_data = get_barang_by_id(barang_id)

                    if barang_data:
                        success, message = add_peminjaman(
                            barang_id,
                            barang_data[1],
                            jumlah_pinjam,
                            tanggal_pinjam,
                            unit,
                            barang_data[3],
                            barang_data[4]
                        )

                        if success:
                            st.session_state.submission_success = True
                            st.session_state.form_submitted = False
                            st.rerun()
                        else:
                            st.error(f"❌ {message}")
                            st.session_state.form_submitted = False
                    else:
                        st.error("❌ Barang tidak ditemukan!")
                        st.session_state.form_submitted = False

                elif submitted:
                    st.info("⏳ Penggunaan sedang diproses...")

        else:
            st.warning("⚠️ Tidak ada barang yang tersedia untuk digunakan.")

    with tab2:
        st.subheader("📜 Riwayat Penggunaan")

        tab_view, tab_delete = st.tabs(["👁️ Lihat Riwayat", "🗑️ Hapus Riwayat"])

        with tab_view:
            df_peminjaman = get_peminjaman()

            if not df_peminjaman.empty:
                col1, col2, col3 = st.columns(3)
                with col1:
                    start_date = st.date_input("📅 Dari Tanggal", value=datetime.now().date() - timedelta(days=30))
                with col2:
                    end_date = st.date_input("📅 Sampai Tanggal", value=datetime.now().date())
                with col3:
                    search_barang = st.text_input("🔍 Cari Barang")

                mask = (pd.to_datetime(df_peminjaman['tanggal_pinjam'], errors='coerce').dt.date >= start_date) & (pd.to_datetime(df_peminjaman['tanggal_pinjam'], errors='coerce').dt.date <= end_date)
                df_filtered = df_peminjaman.loc[mask]

                if search_barang:
                    df_filtered = df_filtered[df_filtered['nama_barang'].str.contains(search_barang, case=False)]

                if not df_filtered.empty:
                    st.info(f"📊 Menampilkan {len(df_filtered)} transaksi penggunaan")

                    display_df = df_filtered.copy()
                    display_df = display_df.rename(columns={
                        'id': 'ID',
                        'nama_barang': 'Nama Barang',
                        'jumlah_pinjam': 'Jumlah Penggunaan',
                        'tanggal_pinjam': 'Tanggal Penggunaan',
                        'unit': 'Unit',
                        'besaran_stok': 'Satuan',
                        'gudang': 'Gudang'
                    })

                    st.dataframe(display_df[['ID', 'Nama Barang', 'Jumlah Penggunaan', 'Tanggal Penggunaan', 'Unit', 'Satuan', 'Gudang']], use_container_width=True)

                    total_transaksi = len(df_filtered)
                    total_barang = df_filtered['jumlah_pinjam'].sum()

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("📊 Total Transaksi", total_transaksi)
                    with col2:
                        st.metric("📦 Total Barang Digunakan", total_barang)

                    # Download Excel
                    create_excel_download(display_df[['ID', 'Nama Barang', 'Jumlah Penggunaan', 'Tanggal Penggunaan', 'Unit', 'Satuan', 'Gudang']], "riwayat_penggunaan", "📥 Download Excel")
                else:
                    st.info("🔭 Tidak ada penggunaan dalam rentang tanggal tersebut.")
            else:
                st.info("🔭 Belum ada riwayat penggunaan.")

        with tab_delete:
            st.warning("⚠️ **PERHATIAN:** Hapus riwayat hanya untuk koreksi kesalahan input. Stok barang TIDAK akan dikembalikan!")

            df_peminjaman = get_peminjaman()

            if not df_peminjaman.empty:
                penggunaan_options = {f"ID-{row['id']}: {row['nama_barang']} ({row['jumlah_pinjam']} {row['besaran_stok']}) - Unit {row['unit']} - {row['tanggal_pinjam']}": row['id']
                                     for _, row in df_peminjaman.iterrows()}

                with st.form("form_hapus_penggunaan"):
                    selected_penggunaan = st.selectbox("🗑️ Pilih riwayat penggunaan yang akan dihapus", list(penggunaan_options.keys()))

                    st.markdown("**Konfirmasi penghapusan:**")
                    confirm = st.checkbox("✅ Saya yakin ingin menghapus riwayat ini")

                    submitted = st.form_submit_button("🗑️ HAPUS RIWAYAT", type="secondary", use_container_width=True)

                    if submitted:
                        if confirm:
                            penggunaan_id = penggunaan_options[selected_penggunaan]
                            success, message = delete_penggunaan(penggunaan_id)

                            if success:
                                st.success(f"✅ {message}")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(f"❌ {message}")
                        else:
                            st.error("❌ Harap centang konfirmasi untuk menghapus riwayat!")
            else:
                st.info("🔭 Belum ada riwayat untuk dihapus.")

# Laporan
elif menu == "📊 Laporan":
    st.header("📊 Laporan Penggunaan")

    df_peminjaman = get_peminjaman()

    if not df_peminjaman.empty:
        df_peminjaman['tanggal_pinjam'] = pd.to_datetime(df_peminjaman['tanggal_pinjam'], errors='coerce')

        st.sidebar.subheader("🏠 Filter Unit")
        unit_options = ["Semua Unit"] + sorted(df_peminjaman['unit'].unique().tolist())
        selected_unit = st.sidebar.selectbox("Pilih Unit", unit_options)

        if selected_unit != "Semua Unit":
            df_peminjaman = df_peminjaman[df_peminjaman['unit'] == selected_unit]
            st.info(f"📋 Menampilkan data untuk unit: {selected_unit}")

        st.subheader("📅 Laporan Harian")
        tanggal_pilih = st.date_input("📅 Pilih Tanggal", value=datetime.now().date())

        df_harian = df_peminjaman[df_peminjaman['tanggal_pinjam'].dt.date == tanggal_pilih]

        if not df_harian.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.metric("📊 Transaksi Hari Ini", len(df_harian))
            with col2:
                st.metric("📦 Total Barang", df_harian['jumlah_pinjam'].sum())

            display_harian = df_harian.copy()
            display_harian['tanggal_pinjam'] = display_harian['tanggal_pinjam'].dt.date
            display_harian = display_harian.rename(columns={
                'nama_barang': 'Nama Barang',
                'jumlah_pinjam': 'Jumlah Penggunaan',
                'tanggal_pinjam': 'Tanggal Penggunaan',
                'unit': 'Unit',
                'besaran_stok': 'Satuan',
                'gudang': 'Gudang'
            })
            st.dataframe(display_harian[['Nama Barang', 'Jumlah Penggunaan', 'Tanggal Penggunaan', 'Unit', 'Satuan', 'Gudang']], use_container_width=True)

            # Download Excel
            create_excel_download(display_harian[['Nama Barang', 'Jumlah Penggunaan', 'Tanggal Penggunaan', 'Unit', 'Satuan', 'Gudang']], "laporan_harian", "📥 Download Excel")

            chart_data = df_harian.groupby('nama_barang')['jumlah_pinjam'].sum().reset_index()
            chart_data = chart_data.rename(columns={'jumlah_pinjam': 'Jumlah Penggunaan', 'nama_barang': 'Nama Barang'})
            if len(chart_data) > 0:
                fig = px.bar(chart_data, x='Nama Barang', y='Jumlah Penggunaan',
                             title=f"📊 Penggunaan per Barang - {tanggal_pilih}",
                             labels={'Jumlah Penggunaan': 'Jumlah Digunakan', 'Nama Barang': 'Nama Barang'})
                fig.update_layout(xaxis_tickangle=-45, height=400, margin=dict(l=20, r=20, t=40, b=80))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"🔭 Tidak ada penggunaan pada tanggal {tanggal_pilih}")

        st.markdown("---")

        st.subheader("📅 Laporan Mingguan")

        col1, col2, col3 = st.columns(3)
        with col1:
            df_weekly = df_peminjaman.copy()
            available_months = df_weekly['tanggal_pinjam'].dt.to_period('M').unique()
            if len(available_months) > 0:
                month_options = [str(month) for month in sorted(available_months)]
                selected_month = st.selectbox("📅 Pilih Bulan", month_options, index=len(month_options)-1 if month_options else 0)

                year, month = map(int, selected_month.split('-'))

                df_month = df_weekly[
                    (df_weekly['tanggal_pinjam'].dt.year == year) &
                    (df_weekly['tanggal_pinjam'].dt.month == month)
                ]

        with col2:
            week_filter = st.selectbox("📊 Filter Minggu", ["Semua Minggu", "Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"])

        if 'df_month' in locals() and not df_month.empty:
            df_month = df_month.copy()
            df_month['iso_week'] = df_month['tanggal_pinjam'].dt.isocalendar().week
            df_month['week_of_month'] = ((df_month['tanggal_pinjam'].dt.day - 1) // 7) + 1
            df_month['minggu'] = df_month['tanggal_pinjam'].apply(
                lambda x: f"{x.year}-W{x.isocalendar()[1]:02d}"
            )

            if week_filter != "Semua Minggu":
                week_num = int(week_filter.split()[1])
                df_month = df_month[df_month['week_of_month'] == week_num]

            weekly_data = df_month.groupby(['minggu', 'nama_barang', 'besaran_stok'])['jumlah_pinjam'].sum().reset_index()

            weekly_data = weekly_data.rename(columns={
                'minggu': 'Minggu',
                'nama_barang': 'Nama Barang',
                'jumlah_pinjam': 'Jumlah Penggunaan',
                'besaran_stok' : 'Satuan'
            })

            if not weekly_data.empty:
                filter_text = f" - {week_filter}" if week_filter != "Semua Minggu" else ""
                st.info(f"📊 Laporan mingguan untuk {selected_month}{filter_text}")
                st.dataframe(weekly_data, use_container_width=True)

                # Download Excel
                create_excel_download(weekly_data, "laporan_mingguan", "📥 Download Excel")

                # CHART BARU - Style seperti Dashboard
                fig_weekly = px.bar(weekly_data, x='Nama Barang', y='Jumlah Penggunaan',
                                    color='Minggu',
                                    title=f"📈 Trend Penggunaan Mingguan - {selected_month}{filter_text}",
                                    labels={'Jumlah Penggunaan': 'Jumlah Digunakan', 'Nama Barang': 'Nama Barang'},
                                    barmode='group')
                fig_weekly.update_layout(
                    xaxis_tickangle=-45,
                    height=400,
                    margin=dict(l=20, r=20, t=40, b=80),
                    showlegend=True
                )
                st.plotly_chart(fig_weekly, use_container_width=True)
            else:
                st.info(f"🔭 Tidak ada data penggunaan untuk {week_filter} bulan {selected_month}")
        elif 'available_months' in locals() and len(available_months) == 0:
            st.info("🔭 Belum ada data penggunaan untuk laporan mingguan")

        st.markdown("---")

        st.subheader("📅 Laporan Bulanan")

        df_monthly = df_peminjaman.copy()
        df_monthly['bulan'] = df_monthly['tanggal_pinjam'].dt.strftime('%Y-%m')
        available_months_monthly = sorted(df_monthly['bulan'].unique().tolist())

        if available_months_monthly:
            col1, col2 = st.columns([1, 2])
            with col1:
                selected_month_filter = st.selectbox(
                    "📅 Pilih Bulan untuk Laporan",
                    ["Semua Bulan"] + available_months_monthly,
                    index=0
                )

            if selected_month_filter != "Semua Bulan":
                df_monthly_filtered = df_monthly[df_monthly['bulan'] == selected_month_filter]
                monthly_data = df_monthly_filtered.groupby(['bulan', 'nama_barang', 'besaran_stok'])['jumlah_pinjam'].sum().reset_index()
                monthly_data = monthly_data.rename(columns={
                    'bulan': 'Bulan',
                    'nama_barang': 'Nama Barang',
                    'jumlah_pinjam': 'Jumlah Penggunaan',
                    'besaran_stok' : 'Satuan'
                })
                chart_title = f"📈 Penggunaan Bulanan - {selected_month_filter}"
            else:
                monthly_data = df_monthly.groupby(['bulan', 'nama_barang', 'besaran_stok'])['jumlah_pinjam'].sum().reset_index()
                monthly_data = monthly_data.rename(columns={
                    'bulan': 'Bulan',
                    'nama_barang': 'Nama Barang',
                    'jumlah_pinjam': 'Jumlah Penggunaan',
                    'besaran_stok' : 'Satuan'
                })
                chart_title = "📈 Trend Penggunaan Bulanan (Semua Bulan)"

            if not monthly_data.empty:
                st.dataframe(monthly_data, use_container_width=True)

                # Download Excel
                create_excel_download(monthly_data, "laporan_bulanan", "📥 Download Excel")

                # CHART BARU - Style seperti Dashboard
                fig_monthly = px.bar(monthly_data, x='Nama Barang', y='Jumlah Penggunaan',
                                     color='Bulan',
                                     title=chart_title,
                                     labels={'Jumlah Penggunaan': 'Jumlah Digunakan', 'Nama Barang': 'Nama Barang'},
                                     barmode='group')
                fig_monthly.update_layout(
                    xaxis_tickangle=-45,
                    height=400,
                    margin=dict(l=20, r=20, t=40, b=80),
                    showlegend=True
                )
                st.plotly_chart(fig_monthly, use_container_width=True)

                if selected_month_filter != "Semua Bulan":
                    total_penggunaan = monthly_data['Jumlah Penggunaan'].sum()
                    total_jenis = len(monthly_data['Nama Barang'].unique())

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("📦 Total Penggunaan", total_penggunaan)
                    with col2:
                        st.metric("📋 Jenis Barang", total_jenis)

    else:
        st.info("🔭 Belum ada data penggunaan untuk membuat laporan.")

# Kelola HPP
elif menu == "💰 Kelola HPP":
    st.header("💰 Kelola HPP (Harga Pokok Produksi)")

    tab1, tab2, tab3 = st.tabs(["➕ Input HPP Manual", "📥 Import HPP dari Excel", "🗑️ Hapus Data HPP"])

    with tab1:
        st.subheader("➕ Input Data HPP Manual")

        with st.form("form_input_hpp", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                unit_options = generate_unit_options()
                unit_hpp = st.text_input("🏠 Unit", placeholder='gudang barat')
                tanggal_hpp = st.date_input("📅 Tanggal", value=datetime.now().date())
                material_hpp = st.text_input("🔨 Nama Material")
            
            with col2:
                harga_hpp = st.number_input("💵 Harga (Rp)", min_value=0, value=0, step=1000)
                keterangan_hpp = st.text_area("📝 Keterangan (Optional)", height=100)

            submitted = st.form_submit_button("➕ Tambah Data HPP", use_container_width=True)

            if submitted:
                if material_hpp.strip() and harga_hpp > 0:
                    add_hpp_data(unit_hpp, tanggal_hpp, material_hpp.strip(), harga_hpp, keterangan_hpp.strip())
                    st.success(f"✅ Data HPP untuk {material_hpp} berhasil ditambahkan!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("❌ Material dan harga harus diisi dengan benar!")

        with tab2:
            st.subheader("📥 Import Data HPP dari Excel")
            st.info("📋 Format Excel: Sheet 'Pengeluaran Material' dengan kolom Tanggal, Material, Unit, Harga")

            uploaded_file_hpp = st.file_uploader("Upload File Excel HPP", type=['xlsx', 'xls'], key="upload_hpp")

            if uploaded_file_hpp is not None:
                try:
                    excel_file = pd.ExcelFile(uploaded_file_hpp)
                    sheet_names = excel_file.sheet_names
                    st.success(f"✅ File berhasil diupload! Ditemukan {len(sheet_names)} sheet.")

                    selected_sheet = st.selectbox("📋 Pilih Sheet", sheet_names)

                    # Baca dan preview data
                    df_preview, total_preview = read_pengeluaran_material(uploaded_file_hpp, sheet_name=selected_sheet, verbose=False)

                    if not df_preview.empty:
                        st.write("**Preview Data (10 baris terakhir):**")
                        st.dataframe(df_preview.tail(10), use_container_width=True)
                        st.metric("💰 Total HPP", f"Rp {total_preview:,.0f}".replace(",", "."))

                        col1, col2 = st.columns(2)
                        with col1:
                            unit_for_import = st.selectbox("🏠 Unit untuk data ini", generate_unit_options(), key="unit_import_hpp")
                        with col2:
                            keterangan_import = st.text_input("📝 Keterangan (Optional)", key="ket_import_hpp")

                        if st.button("🚀 Import Data HPP", type="primary", use_container_width=True):
                            with st.spinner("Memproses import..."):
                                conn = sqlite3.connect('inventory_rumah.db')
                                c = conn.cursor()

                                imported_count = 0
                                for _, row in df_preview.iterrows():
                                    if pd.notna(row['Tanggal']) and pd.notna(row['Harga']):
                                        tanggal_val = row['Tanggal']
                                        # jika Timestamp atau datetime -> format ke DD/MM/YYYY
                                        if isinstance(tanggal_val, (pd.Timestamp, datetime)):
                                            tanggal_str = tanggal_val.strftime('%d/%m/%Y')
                                        elif isinstance(tanggal_val, str):
                                            # possible raw strings: coba bersihkan lalu parse / normalisasi
                                            t = tanggal_val.strip()
                                            # jika format sudah YYYY-MM-DD, convert dulu ke dd/mm/YYYY
                                            try:
                                                parsed = pd.to_datetime(t, errors='coerce', dayfirst=False)
                                                if pd.notna(parsed):
                                                    # format menjadi DD/MM/YYYY
                                                    tanggal_str = parsed.strftime('%d/%m/%Y')
                                                else:
                                                    tanggal_str = t  # fallback: simpan apa adanya
                                            except:
                                                tanggal_str = t
                                        else:
                                            tanggal_str = str(tanggal_val)

                                        c.execute("""INSERT INTO hpp (unit, tanggal, material, harga, keterangan)
                                                    VALUES (?, ?, ?, ?, ?)""",
                                                (unit_for_import, tanggal_str, row['Material'], row['Harga'], keterangan_import))
                                        imported_count += 1

                                conn.commit()
                                conn.close()

                                st.success(f"✅ Berhasil import {imported_count} data HPP!")
                                st.balloons()
                                time.sleep(2)
                                st.rerun()

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")


    with tab3:
        st.subheader("🗑️ Hapus Data HPP")
        st.warning("⚠️ Penghapusan data HPP bersifat permanen!")

        df_hpp = get_hpp_data()

        if not df_hpp.empty:
            col1, col2 = st.columns(2)
            with col1:
                filter_unit_delete = st.selectbox("🏠 Filter Unit", ["Semua"] + generate_unit_options(), key="delete_unit_filter")
            with col2:
                search_material_delete = st.text_input("🔍 Cari Material", key="delete_material_search")

            df_filtered = df_hpp.copy()
            if filter_unit_delete != "Semua":
                df_filtered = df_filtered[df_filtered['unit'] == filter_unit_delete]
            if search_material_delete:
                df_filtered = df_filtered[df_filtered['material'].str.contains(search_material_delete, case=False, na=False)]

            if not df_filtered.empty:
                hpp_options = {f"ID-{row['id']}: {row['unit']} - {row['material']} - Rp {row['harga']:,.0f} ({row['tanggal']})": row['id']
                               for _, row in df_filtered.iterrows()}

                with st.form("form_hapus_hpp"):
                    selected_hpp = st.selectbox("🗑️ Pilih data yang akan dihapus", list(hpp_options.keys()))

                    confirm = st.checkbox("✅ Saya yakin ingin menghapus data ini")

                    submitted = st.form_submit_button("🗑️ HAPUS DATA", type="secondary", use_container_width=True)

                    if submitted:
                        if confirm:
                            hpp_id = hpp_options[selected_hpp]
                            success, message = delete_hpp(hpp_id)

                            if success:
                                st.success(f"✅ {message}")
                                time.sleep(1)
                                st.rerun()
                        else:
                            st.error("❌ Harap centang konfirmasi untuk menghapus data!")
            else:
                st.info("🔭 Tidak ada data HPP yang sesuai filter.")
        else:
            st.info("🔭 Belum ada data HPP.")

# Laporan HPP
elif menu == "💰 Laporan HPP":
    st.header("💰 Laporan HPP (Harga Pokok Produksi)")

    tab1, tab2, tab3 = st.tabs(["📊 Laporan per Unit", "📈 Laporan Periode", "📋 Ringkasan Total"])

    with tab1:
        st.subheader("📊 Laporan HPP per Unit")

        col1, col2, col3 = st.columns(3)
        with col1:
            unit_options = ["Semua Unit"] + generate_unit_options()
            selected_unit_hpp = st.selectbox(
                "🏠 Pilih Unit",
                unit_options,
                key="filter_unit_hpp"
            )
        with col2:
            start_date_hpp = st.date_input(
                "📅 Dari Tanggal",
                value=datetime.now().date() - timedelta(days=30),
                key="hpp_start"
            )
        with col3:
            end_date_hpp = st.date_input(
                "📅 Sampai Tanggal",
                value=datetime.now().date(),
                key="hpp_end"
            )

        # --- Ambil data sesuai filter ---
        df_hpp_filtered = get_hpp_data(
            unit=None if selected_unit_hpp == "Semua Unit" else selected_unit_hpp,
            start_date=start_date_hpp,
            end_date=end_date_hpp
        )

        if not df_hpp_filtered.empty:
            df_hpp_filtered['tanggal'] = pd.to_datetime(
                df_hpp_filtered['tanggal'], errors='coerce', dayfirst=True
            ).dt.strftime('%d/%m/%Y')

            total_hpp = df_hpp_filtered['harga'].sum()
            jumlah_transaksi = len(df_hpp_filtered)

            col1, col2 = st.columns(2)
            with col1:
                st.metric("💰 Total HPP", f"Rp {total_hpp:,.0f}".replace(",", "."))
            with col2:
                st.metric("📝 Jumlah Transaksi", jumlah_transaksi)

            st.dataframe(
                df_hpp_filtered[['id', 'unit', 'tanggal', 'material', 'harga', 'keterangan']],
                width="stretch"
            )
            create_excel_download(
                df_hpp_filtered[['id', 'unit', 'tanggal', 'material', 'harga', 'keterangan']],
                "laporan_hpp_unit",
                "📥 Download Excel"
            )
            chart_data = (
                df_hpp_filtered.groupby('material')['harga']
                .sum()
                .reset_index()
                .sort_values('harga', ascending=False)
                .head(10)
            )
            fig = px.bar(
                chart_data,
                x='material',
                y='harga',
                title=f"Top 10 Material Termahal - {selected_unit_hpp}",
                labels={'harga': 'Total Harga (Rp)', 'material': 'Material'}
            )
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("🔭 Tidak ada data HPP untuk filter yang dipilih.")

    with tab2:
        st.subheader("📈 Laporan HPP Periode")

        df_hpp = get_hpp_data()

        if not df_hpp.empty:
            # konversi tanggal ke datetime dayfirst agar sesuai format dd/mm/yyyy
            df_hpp['tanggal'] = pd.to_datetime(
                df_hpp['tanggal'], errors='coerce', dayfirst=True
            )

            # Filter Unit
            filter_unit = st.selectbox(
                "🏠 Pilih Unit",
                ["Semua"] + sorted(df_hpp['unit'].dropna().unique().tolist()),
                key="filter_unit_periode"
            )
            if filter_unit != "Semua":
                df_hpp = df_hpp[df_hpp['unit'] == filter_unit]

            # Tambah kolom bulan (format Month/Year)
            df_hpp['bulan'] = df_hpp['tanggal'].dt.strftime('%m/%Y')

            # Laporan Bulanan
            st.markdown("### 📅 Laporan Bulanan")
            monthly_data = (
                df_hpp.groupby(['bulan', 'unit'])['harga']
                .sum()
                .reset_index()
                .rename(columns={'bulan': 'Bulan', 'unit': 'Unit', 'harga': 'Total HPP'})
            )

            if not monthly_data.empty:
                st.dataframe(monthly_data, width="stretch")

                # Download Excel
                create_excel_download(
                    monthly_data,
                    "laporan_hpp_bulanan",
                    "📥 Download Excel"
                )

                fig_monthly = px.bar(
                    monthly_data,
                    x='Unit',
                    y='Total HPP',
                    color='Bulan',
                    title="Trend HPP Bulanan per Unit",
                    labels={'Total HPP': 'Total HPP (Rp)', 'Unit': 'Unit'},
                    barmode='group'
                )
                st.plotly_chart(fig_monthly, use_container_width=True)

            # Ringkasan per Unit
            st.markdown("### 🏠 Ringkasan per Unit")
            unit_summary = (
                df_hpp.groupby('unit')['harga']
                .agg(['sum', 'count', 'mean'])
                .reset_index()
                .rename(
                    columns={
                        'unit': 'Unit',
                        'sum': 'Total HPP',
                        'count': 'Jumlah Transaksi',
                        'mean': 'Rata-rata HPP'
                    }
                )
                .sort_values('Total HPP', ascending=False)
            )

            st.dataframe(unit_summary, width="stretch")

            # Download Excel
            create_excel_download(
                unit_summary,
                "ringkasan_hpp_unit",
                "📥 Download Excel"
            )

            # Pie Chart
            fig_pie = px.pie(
                unit_summary,
                values='Total HPP',
                names='Unit',
                title="Distribusi HPP per Unit"
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("🔭 Belum ada data HPP.")

    with tab3:
        st.subheader("📋 Ringkasan Total")

        df_hpp = get_hpp_data()

        if not df_hpp.empty:
            # Konversi tanggal string ke datetime untuk keperluan analisis
            df_hpp['tanggal'] = pd.to_datetime(df_hpp['tanggal'], errors='coerce', dayfirst=True)

            # Filter Unit
            filter_unit_total = st.selectbox(
                "🏠 Pilih Unit",
                ["Semua"] + sorted(df_hpp['unit'].dropna().unique().tolist()),
                key="filter_unit_total"
            )

            if filter_unit_total != "Semua":
                df_hpp = df_hpp[df_hpp['unit'] == filter_unit_total]

            # Pastikan tanggal ditampilkan kembali dalam format dd/mm/yyyy
            df_hpp['tanggal'] = df_hpp['tanggal'].dt.strftime('%d/%m/%Y')

            # Hitung metrik ringkasan
            total_hpp_all = df_hpp['harga'].sum()
            total_transaksi = len(df_hpp)
            rata_rata_hpp = df_hpp['harga'].mean()
            hpp_tertinggi = df_hpp['harga'].max()
            hpp_terendah = df_hpp['harga'].min()

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("💰 Total HPP Keseluruhan", f"Rp {total_hpp_all:,.0f}".replace(",", "."))
                st.metric("📝 Total Transaksi", total_transaksi)
            with col2:
                st.metric("📊 Rata-rata HPP", f"Rp {rata_rata_hpp:,.0f}".replace(",", "."))
                st.metric("⬆️ HPP Tertinggi", f"Rp {hpp_tertinggi:,.0f}".replace(",", "."))
            with col3:
                st.metric("⬇️ HPP Terendah", f"Rp {hpp_terendah:,.0f}".replace(",", "."))

            # 🔝 Top Material
            st.markdown("### 🔝 Top 15 Material dengan HPP Tertinggi")
            top_materials = (
                df_hpp.groupby('material')['harga']
                .sum()
                .reset_index()
                .sort_values('harga', ascending=False)
                .head(15)
            )
            top_materials['harga_formatted'] = top_materials['harga'].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))

            st.dataframe(top_materials[['material', 'harga_formatted']], width="stretch")

            # Bar Chart
            fig_top = px.bar(
                top_materials,
                x='material',
                y='harga',
                title=f"Top 15 Material dengan Total HPP Tertinggi ({filter_unit_total})",
                labels={'harga': 'Total HPP (Rp)', 'material': 'Material'}
            )
            fig_top.update_layout(xaxis_tickangle=-45, height=500)
            st.plotly_chart(fig_top, use_container_width=True)

            # Download Excel
            create_excel_download(
                top_materials[['material', 'harga_formatted']],
                "top15_material_hpp",
                "📥 Download Excel"
            )

        else:
            st.info("🔭 Belum ada data HPP.")

# Stok Rendah
elif menu == "⚠️ Stok Rendah":
    st.header("⚠️ Monitor Stok Rendah")

    stok_rendah = check_stok_rendah()

    if not stok_rendah.empty:
        st.error(f"🚨 PERINGATAN! Ada {len(stok_rendah)} barang dengan stok kurang dari 20!")

        for _, item in stok_rendah.iterrows():
            with st.container():
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f"**📦 {item['nama_barang']}**")
                with col2:
                    st.markdown(f"**📊 Stok:** {item['stok']} {item['besaran_stok']}")
                with col3:
                    st.markdown(f"**🏭 Gudang:** {item['gudang']}")
                st.markdown("---")

        fig = px.bar(stok_rendah, x='nama_barang', y='stok',
                     color='gudang', title="📊 Barang dengan Stok Rendah")
        fig.add_hline(y=20, line_dash="dash", line_color="red", annotation_text="⚠️ Batas Minimum (20)")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("💡 Saran Restock")
        for _, item in stok_rendah.iterrows():
            saran_restock = max(50, item['stok'] * 3)
            st.info(f"**{item['nama_barang']}**: Disarankan menambah stok hingga **{saran_restock} {item['besaran_stok']}**")
    else:
        st.success("🎉 Semua barang memiliki stok yang mencukupi!")

        df_barang = get_barang()
        if not df_barang.empty:
            df_aman = df_barang[df_barang['stok'] >= 20].nsmallest(5, 'stok')
            if not df_aman.empty:
                st.subheader("📊 5 Barang dengan Stok Terendah (Masih Aman)")
                st.dataframe(df_aman[['nama_barang', 'stok', 'besaran_stok', 'gudang']], use_container_width=True)

# Import/Export Data
elif menu == "📥 Import/Export Data":
    st.header("📥 Import/Export Data")

    tab1, tab2, tab3 = st.tabs(["📦 Import Data Barang", "📤 Import Riwayat Penggunaan", "💾 Export/Backup Data"])

    # TAB 1: IMPORT DATA BARANG (BARU)
    with tab1:
        st.subheader("📦 Import Data Barang Masuk dari Excel")
        st.warning("⚠️ **Format Excel Multi-Baris:** Header dibaca dari Baris 3 (Nama Barang, Jumlah, Satuan) dan Baris 4 (Sen-Min)")
        st.info("📋 Barang yang diimport akan ditambahkan ke stok yang sudah ada berdasarkan hari barang masuk.")

        st.markdown("""
        **Catatan Penting:**
        - **Kolom A**: Diabaikan (biasanya untuk nomor urut)
        - **Baris 3**: Mulai dari **Kolom B** → NAMA BARANG, JUMLAH (diabaikan), SATUAN
        - **Baris 4**: SEN, SEL, RAB, KAM, JUM, SAB, MIN (hari barang masuk)
        - Data dimulai dari **Baris 5**
        - Pilih **tanggal Senin** untuk setiap sheet
        - Stok barang akan **ditambahkan** sesuai hari barang masuk
        """)

        uploaded_file_barang = st.file_uploader("Upload File Excel untuk Data Barang Masuk", type=['xlsx', 'xls'], key="upload_barang")

        if uploaded_file_barang is not None:
            try:
                excel_file = pd.ExcelFile(uploaded_file_barang)
                sheet_names = excel_file.sheet_names

                st.success(f"✅ File berhasil diupload! Ditemukan {len(sheet_names)} sheet.")

                for name in sheet_names:
                    if name not in st.session_state.selected_sheets_barang:
                        st.session_state.selected_sheets_barang[name] = True

                st.markdown("---")
                st.subheader("🔧 Konfigurasi Import per Sheet")

                selected_sheets_barang = []

                for sheet_name in sheet_names:
                    default_date = st.session_state.import_barang_config.get(sheet_name, {}).get('tanggal_senin', datetime.now().date())
                    default_gudang = st.session_state.import_barang_config.get(sheet_name, {}).get('gudang', 'Gudang 1')

                    is_selected = st.checkbox(f"✅ Pilih Sheet: **{sheet_name}**",
                                              value=st.session_state.selected_sheets_barang[sheet_name],
                                              key=f"check_barang_{sheet_name}")
                    st.session_state.selected_sheets_barang[sheet_name] = is_selected

                    if is_selected:
                        selected_sheets_barang.append(sheet_name)
                        with st.expander(f"⚙️ Konfigurasi untuk {sheet_name}", expanded=False):
                            col1, col2 = st.columns(2)

                            with col1:
                                tanggal_senin = st.date_input(
                                    f"📅 Tanggal Senin minggu ini",
                                    value=default_date,
                                    key=f"date_barang_{sheet_name}",
                                    help="Pilih tanggal hari Senin dari minggu data barang masuk"
                                )

                            with col2:
                                gudang = st.selectbox(
                                    f"🏭 Gudang untuk sheet '{sheet_name}'",
                                    ["Gudang 1", "Gudang 2"],
                                    index=0 if default_gudang == "Gudang 1" else 1,
                                    key=f"gudang_barang_{sheet_name}"
                                )

                            # Preview data dengan multi-header
                            try:
                                df_row3 = pd.read_excel(uploaded_file_barang, sheet_name=sheet_name, header=2, nrows=0)
                                header_row3 = df_row3.columns.tolist()

                                df_row4 = pd.read_excel(uploaded_file_barang, sheet_name=sheet_name, header=3, nrows=0)
                                header_row4 = df_row4.columns.tolist()

                                combined_header = []
                                num_cols = max(len(header_row3), len(header_row4))

                                for i in range(num_cols):
                                    h3 = header_row3[i] if i < len(header_row3) else ''
                                    h4 = header_row4[i] if i < len(header_row4) else ''

                                    h3_clean = str(h3).lower().strip().replace(' ', '')
                                    h4_clean = str(h4).lower().strip().replace(' ', '')

                                    if i == 0:
                                        combined_header.append('skip_a')
                                    elif i == 1:
                                        combined_header.append('namabarang')
                                    elif i == 2:
                                        combined_header.append('skip_jumlah')
                                    elif i == 3:
                                        combined_header.append('satuan')
                                    elif h4_clean in ['sen', 'sel', 'rab', 'kam', 'jum', 'sab', 'min']:
                                        combined_header.append(h4_clean)
                                    else:
                                        combined_header.append('skip_' + str(i))

                                df_preview = pd.read_excel(uploaded_file_barang, sheet_name=sheet_name, header=None, skiprows=4, nrows=5)

                                num_data_cols = len(df_preview.columns)
                                if len(combined_header) >= num_data_cols:
                                    df_preview.columns = combined_header[:num_data_cols]
                                else:
                                    df_preview.columns = combined_header + [f'extra_{j}' for j in range(len(combined_header), num_data_cols)]

                                st.write("**Preview Data (5 Baris Pertama):**")
                                st.dataframe(df_preview.head(5), use_container_width=True)

                            except Exception as e:
                                st.error(f"Error membaca preview: {str(e)}")

                            st.session_state.import_barang_config[sheet_name] = {
                                'tanggal_senin': tanggal_senin,
                                'gudang': gudang
                            }
                    else:
                        if sheet_name in st.session_state.import_barang_config:
                            del st.session_state.import_barang_config[sheet_name]

                st.markdown("---")
                st.info(f"Total {len(selected_sheets_barang)} sheet terpilih untuk diimpor.")

                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🚀 Proses Import Data Barang Masuk", type="primary", use_container_width=True, key="import_barang_btn"):
                        if not selected_sheets_barang:
                            st.error("❌ Tidak ada sheet yang dipilih untuk diimpor!")
                            st.stop()

                        with st.spinner("Memproses import data barang masuk..."):
                            total_imported = 0
                            total_updated = 0
                            errors = []

                            conn = sqlite3.connect('inventory_rumah.db')
                            c = conn.cursor()

                            for sheet_name in selected_sheets_barang:
                                try:
                                    # Baca multi-header
                                    df_row3 = pd.read_excel(uploaded_file_barang, sheet_name=sheet_name, header=2, nrows=0)
                                    header_row3 = df_row3.columns.tolist()

                                    df_row4 = pd.read_excel(uploaded_file_barang, sheet_name=sheet_name, header=3, nrows=0)
                                    header_row4 = df_row4.columns.tolist()

                                    combined_header = []
                                    num_cols = max(len(header_row3), len(header_row4))

                                    for i in range(num_cols):
                                        h3 = header_row3[i] if i < len(header_row3) else ''
                                        h4 = header_row4[i] if i < len(header_row4) else ''

                                        h3_clean = str(h3).lower().strip().replace(' ', '')
                                        h4_clean = str(h4).lower().strip().replace(' ', '')

                                        if i == 0:
                                            combined_header.append('skip_a')
                                        elif i == 1:
                                            combined_header.append('namabarang')
                                        elif i == 2:
                                            combined_header.append('skip_jumlah')
                                        elif i == 3:
                                            combined_header.append('satuan')
                                        elif h4_clean in ['sen', 'sel', 'rab', 'kam', 'jum', 'sab', 'min']:
                                            combined_header.append(h4_clean)
                                        else:
                                            combined_header.append('skip_' + str(i))

                                    df = pd.read_excel(uploaded_file_barang, sheet_name=sheet_name, header=None, skiprows=4)
                                    num_data_cols = len(df.columns)

                                    if len(combined_header) >= num_data_cols:
                                        df.columns = combined_header[:num_data_cols]
                                    else:
                                        df.columns = combined_header + [f'extra_{j}' for j in range(len(combined_header), num_data_cols)]

                                    config = st.session_state.import_barang_config.get(sheet_name)
                                    if not config:
                                        errors.append(f"Sheet '{sheet_name}': Konfigurasi tidak ditemukan.")
                                        continue

                                    tanggal_senin = config['tanggal_senin']
                                    gudang = config['gudang']

                                    hari_cols = ['sen', 'sel', 'rab', 'kam', 'jum', 'sab', 'min']

                                    for idx, row in df.iterrows():
                                        nama_barang = str(row.get('namabarang', '')).strip()
                                        satuan = str(row.get('satuan', '')).strip()

                                        if not nama_barang or nama_barang == 'nan':
                                            continue

                                        if not satuan or satuan == 'nan':
                                            satuan = 'pcs'

                                        for day_idx, hari in enumerate(hari_cols):
                                            if hari not in df.columns:
                                                continue

                                            jumlah = row.get(hari, 0)

                                            try:
                                                if pd.isna(jumlah) or jumlah is None or str(jumlah).strip() == '':
                                                    jumlah = 0
                                                else:
                                                    jumlah = int(float(jumlah))
                                            except Exception:
                                                jumlah = 0

                                            if jumlah <= 0:
                                                continue

                                            tanggal_masuk = tanggal_senin + timedelta(days=day_idx)

                                            c.execute("SELECT id, stok FROM barang WHERE LOWER(nama_barang) = LOWER(?) AND LOWER(gudang) = LOWER(?)",
                                                      (nama_barang, gudang))
                                            existing = c.fetchone()

                                            if existing:
                                                # Update stok yang sudah ada
                                                barang_id, stok_lama = existing
                                                stok_baru = stok_lama + jumlah
                                                c.execute("UPDATE barang SET stok = ?, besaran_stok = ? WHERE id = ?",
                                                          (stok_baru, satuan, barang_id))

                                                # Catat riwayat
                                                c.execute("""INSERT INTO riwayat_stok
                                                            (barang_id, nama_barang, jumlah_tambah, stok_sebelum, stok_sesudah, gudang, tanggal_tambah)
                                                            VALUES (?, ?, ?, ?, ?, ?, ?)""",
                                                            (barang_id, nama_barang, jumlah, stok_lama, stok_baru, gudang, tanggal_masuk))

                                                total_updated += 1
                                            else:
                                                # Tambah barang baru
                                                c.execute("INSERT INTO barang (nama_barang, stok, besaran_stok, gudang, created_at) VALUES (?, ?, ?, ?, ?)",
                                                          (nama_barang, jumlah, satuan, gudang, tanggal_masuk))

                                                barang_id = c.lastrowid

                                                # Catat riwayat
                                                c.execute("""INSERT INTO riwayat_stok
                                                            (barang_id, nama_barang, jumlah_tambah, stok_sebelum, stok_sesudah, gudang, tanggal_tambah)
                                                            VALUES (?, ?, ?, ?, ?, ?, ?)""",
                                                            (barang_id, nama_barang, jumlah, 0, jumlah, gudang, tanggal_masuk))

                                                total_imported += 1

                                except Exception as e:
                                    errors.append(f"Sheet '{sheet_name}': {str(e)}")

                            conn.commit()
                            conn.close()
                            upload_after_write(LOCAL_DB)

                            if total_imported > 0 or total_updated > 0:
                                st.success(f"✅ Berhasil import barang masuk! **{total_imported}** barang baru dan **{total_updated}** penambahan stok!")
                                st.balloons()
                            else:
                                st.warning("⚠️ Tidak ada barang yang berhasil diimport. Cek format file Excel Anda.")

                            if errors:
                                st.error("⚠️ Beberapa **error** terjadi selama import:")
                                for error in errors:
                                    st.write(f"- {error}")

                            st.session_state.import_barang_config = {}
                            st.session_state.selected_sheets_barang = {}
                            time.sleep(2)
                            st.rerun()

            except Exception as e:
                st.error(f"❌ Error membaca file: {str(e)}")
                st.write("Pastikan file Excel Anda memiliki format yang benar dan tidak rusak.")

    # TAB 2: IMPORT RIWAYAT PENGGUNAAN
    with tab2:
        st.subheader("📤 Import Riwayat Penggunaan dari Excel")
        st.warning("⚠️ **Format Excel Multi-Baris:** Header dibaca dari Baris 2 (Nama Barang, Satuan) dan Baris 3 (Sen-Min)")
        st.info("📋 Sistem akan mencoba menggabungkan header dari kedua baris untuk membaca data dengan benar.")

        st.markdown("""
        **Catatan Penting:**
        - **Baris 2**: NAMA BARANG dimulai dari **Kolom B**. Kolom A diabaikan.
        - **Baris 3**: SEN, SEL, RAB, KAM, JUM, SAB, MIN
        - Kolom `JUMLAH` (Qty) di **Kolom C** akan **diabaikan**.
        - Stok barang **TIDAK** akan dikurangi.
        """)

        uploaded_file = st.file_uploader("Upload File Excel", type=['xlsx', 'xls'], key="upload_penggunaan")

        if uploaded_file is not None:
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names

                st.success(f"✅ File berhasil diupload! Ditemukan {len(sheet_names)} sheet.")

                # Inisialisasi state sheet terpilih (WAJIB)
                if 'selected_sheets_for_import' not in st.session_state:
                    st.session_state.selected_sheets_for_import = []

                for name in sheet_names:
                    if name not in st.session_state.selected_sheets:
                        st.session_state.selected_sheets[name] = True

                st.markdown("---")
                st.subheader("🔧 Konfigurasi Import per Sheet")

                st.session_state.selected_sheets_for_import = []

                for sheet_name in sheet_names:
                    default_unit = st.session_state.import_config.get(sheet_name, {}).get('unit', 'A1')
                    default_date = st.session_state.import_config.get(sheet_name, {}).get('tanggal_senin', datetime.now().date())

                    is_selected = st.checkbox(f"✅ Pilih Sheet: **{sheet_name}**",
                                              value=st.session_state.selected_sheets[sheet_name],
                                              key=f"check_{sheet_name}")
                    st.session_state.selected_sheets[sheet_name] = is_selected

                    if is_selected:
                        st.session_state.selected_sheets_for_import.append(sheet_name)
                        with st.expander(f"⚙️ Konfigurasi untuk {sheet_name}", expanded=False):
                            col1, col2 = st.columns(2)

                            with col1:
                                    unit = st.text_input(
                                        f"🏠 Unit untuk sheet '{sheet_name}'",
                                        placeholder="Contoh: TOTAL / Gudang Barat / Proyek A",
                                        key=f"unit_{sheet_name}"
                                    )

                            with col2:
                                tanggal_senin = st.date_input(
                                    f"📅 Tanggal Senin minggu ini",
                                    value=default_date,
                                    key=f"date_{sheet_name}",
                                    help="Pilih tanggal hari Senin dari minggu data ini"
                                )

                            try:
                                df_row2 = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1, nrows=0)
                                header_row2 = df_row2.columns.tolist()

                                df_row3 = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=2, nrows=0)
                                header_row3 = df_row3.columns.tolist()

                                combined_header = []
                                num_cols = max(len(header_row2), len(header_row3))

                                for i in range(num_cols):
                                    h2 = header_row2[i] if i < len(header_row2) else ''
                                    h3 = header_row3[i] if i < len(header_row3) else ''

                                    h2_clean = str(h2).lower().strip().replace(' ', '')
                                    h3_clean = str(h3).lower().strip().replace(' ', '')

                                    if i == 0:
                                        combined_header.append('skip_a')
                                    elif i == 1:
                                        combined_header.append('namabarang')
                                    elif i == 2:
                                        combined_header.append('skip_jumlah')
                                    elif i == 3:
                                        combined_header.append('satuan')
                                    elif h3_clean in ['sen', 'sel', 'rab', 'kam', 'jum', 'sab', 'min']:
                                        combined_header.append(h3_clean)
                                    else:
                                        combined_header.append('skip_' + str(i))

                                df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, skiprows=3, nrows=5)

                                num_data_cols = len(df_preview.columns)
                                if len(combined_header) >= num_data_cols:
                                    df_preview.columns = combined_header[:num_data_cols]
                                else:
                                    df_preview.columns = combined_header + [f'extra_{j}' for j in range(len(combined_header), num_data_cols)]

                                st.write("**Preview Data (5 Baris Pertama):**")
                                st.dataframe(df_preview.head(5), use_container_width=True)

                            except Exception as e:
                                st.error(f"Error membaca preview: {str(e)}")

                            st.session_state.import_config[sheet_name] = {
                                'unit': unit,
                                'tanggal_senin': tanggal_senin
                            }
                    else:
                        if sheet_name in st.session_state.import_config:
                            del st.session_state.import_config[sheet_name]

                st.markdown("---")
                st.info(f"Total {len(st.session_state.selected_sheets_for_import)} sheet terpilih untuk diimpor.")

                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🚀 Proses Import Sheet Terpilih", type="primary", use_container_width=True, key="import_penggunaan_btn"):
                        if not st.session_state.selected_sheets_for_import:
                            st.error("❌ Tidak ada sheet yang dipilih untuk diimpor!")
                            st.stop()

                        with st.spinner("Memproses import data..."):
                            total_imported = 0
                            errors = []

                            conn = sqlite3.connect('inventory_rumah.db')
                            c = conn.cursor()

                            for sheet_name in st.session_state.selected_sheets_for_import:
                                config = st.session_state.import_config.get(sheet_name)
                                if not config or not config.get('unit', '').strip():
                                    st.error(f"❌ Unit untuk sheet '{sheet_name}' wajib diisi!")
                                    st.stop()
                                try:
                                    df_row2 = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1, nrows=0)
                                    header_row2 = df_row2.columns.tolist()

                                    df_row3 = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=2, nrows=0)
                                    header_row3 = df_row3.columns.tolist()

                                    combined_header = []
                                    num_cols = max(len(header_row2), len(header_row3))

                                    for i in range(num_cols):
                                        h2 = header_row2[i] if i < len(header_row2) else ''
                                        h3 = header_row3[i] if i < len(header_row3) else ''

                                        h2_clean = str(h2).lower().strip().replace(' ', '')
                                        h3_clean = str(h3).lower().strip().replace(' ', '')

                                        if i == 0:
                                            combined_header.append('skip_a')
                                        elif i == 1:
                                            combined_header.append('namabarang')
                                        elif i == 2:
                                            combined_header.append('skip_jumlah')
                                        elif i == 3:
                                            combined_header.append('satuan')
                                        elif h3_clean in ['sen', 'sel', 'rab', 'kam', 'jum', 'sab', 'min']:
                                            combined_header.append(h3_clean)
                                        else:
                                            combined_header.append('skip_' + str(i))

                                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, skiprows=3)
                                    num_data_cols = len(df.columns)

                                    if len(combined_header) >= num_data_cols:
                                        df.columns = combined_header[:num_data_cols]
                                    else:
                                        df.columns = combined_header + [f'extra_{j}' for j in range(len(combined_header), num_data_cols)]

                                    config = st.session_state.import_config.get(sheet_name)
                                    if not config:
                                        errors.append(f"Sheet '{sheet_name}': Konfigurasi tidak ditemukan.")
                                        continue

                                    unit = config['unit']
                                    tanggal_senin = config['tanggal_senin']

                                    hari_cols = ['sen', 'sel', 'rab', 'kam', 'jum', 'sab', 'min']

                                    for idx, row in df.iterrows():
                                        nama_barang = str(row.get('namabarang', '')).strip()
                                        satuan = str(row.get('satuan', '')).strip()

                                        if not nama_barang or nama_barang == 'nan':
                                            continue

                                        if not satuan or satuan == 'nan':
                                            satuan = 'pcs'

                                        for day_idx, hari in enumerate(hari_cols):
                                            if hari not in df.columns:
                                                continue

                                            jumlah = row.get(hari, 0)

                                            try:
                                                if pd.isna(jumlah) or jumlah is None or str(jumlah).strip() == '':
                                                    jumlah = 0
                                                else:
                                                    jumlah = int(float(jumlah))
                                            except Exception:
                                                jumlah = 0

                                            if jumlah <= 0:
                                                continue

                                            tanggal_penggunaan = tanggal_senin + timedelta(days=day_idx)

                                            # --- Tambah ke tabel peminjaman seperti biasa ---
                                            c.execute("""INSERT INTO peminjaman
                                                            (barang_id, nama_barang, jumlah_pinjam, tanggal_pinjam,
                                                            unit, besaran_stok, gudang)
                                                            VALUES (NULL, ?, ?, ?, ?, ?, 'Gudang 1')""",
                                                        (nama_barang, jumlah, tanggal_penggunaan, unit, satuan))

                                            # --- Kurangi stok langsung berdasarkan nama barang ---
                                            c.execute("SELECT id, stok, gudang FROM barang WHERE LOWER(nama_barang) = LOWER(?)", (nama_barang,))
                                            barang_data = c.fetchone()

                                            if barang_data:
                                                barang_id, stok_sekarang, gudang = barang_data
                                                stok_baru = stok_sekarang - jumlah
                                                if stok_baru < 0:
                                                    stok_baru = 0

                                                c.execute("UPDATE barang SET stok = ? WHERE id = ?", (stok_baru, barang_id))

                                                # Tambahkan riwayat stok sebagai pengurangan
                                                c.execute("""
                                                    INSERT INTO riwayat_stok
                                                    (barang_id, nama_barang, jumlah_tambah, stok_sebelum, stok_sesudah, gudang, tanggal_tambah)
                                                    VALUES (?, ?, ?, ?, ?, ?, ?)
                                                """, (barang_id, nama_barang, -jumlah, stok_sekarang, stok_baru, gudang, tanggal_penggunaan))

                                            total_imported += 1

                                except Exception as e:
                                    errors.append(f"Sheet '{sheet_name}': {str(e)}")

                            conn.commit()
                            conn.close()

                            if total_imported > 0:
                                st.success(f"✅ Berhasil import **{total_imported}** transaksi penggunaan dari {len(st.session_state.selected_sheets_for_import)} sheet!")
                                st.balloons()
                            else:
                                st.warning("⚠️ Tidak ada transaksi yang berhasil diimport. Cek format file Excel Anda.")

                            if errors:
                                st.error("⚠️ Beberapa **error** terjadi selama import:")
                                for error in errors:
                                    st.write(f"- {error}")

                            st.session_state.import_config = {}
                            st.session_state.selected_sheets = {}
                            time.sleep(2)
                            st.rerun()

            except Exception as e:
                st.error(f"❌ Error membaca file: {str(e)}")
                st.write("Pastikan file Excel Anda memiliki format yang benar dan tidak rusak.")
    with tab3:
        st.subheader("💾 Export/Backup Data")
        st.markdown("---")

        st.markdown("""
        **Export data aplikasi ke Excel untuk backup atau analisis lebih lanjut.**

        File akan berisi 3 sheet:
        1. **Data Barang**
        2. **Riwayat Penggunaan**
        3. **Riwayat Kelola Barang**
        """)

        col1, col2 = st.columns(2)

        with col1:
            if st.button("📥 Buat File Backup Excel", type="primary", use_container_width=True):
                st.session_state['ready_to_download_excel'] = True

            if st.session_state.get('ready_to_download_excel'):
                try:
                    df_barang = get_barang()
                    df_penggunaan = get_peminjaman()
                    df_riwayat = get_riwayat_stok()

                    df_penggunaan_export = df_penggunaan.copy().rename(columns={
                        'tanggal_pinjam': 'Tanggal Penggunaan'
                    })

                    df_riwayat_export = df_riwayat.copy().rename(columns={
                        'jumlah_tambah': 'Jumlah Perubahan (+/-)',
                        'tanggal_tambah': 'Tanggal Transaksi',
                        'stok_sebelum': 'Stok Sebelum',
                        'stok_sesudah': 'Stok Sesudah',
                    })

                    sheets_to_export = [
                        ('Data Barang', df_barang),
                        ('Riwayat Penggunaan', df_penggunaan_export),
                        ('Riwayat Kelola Barang', df_riwayat_export)
                    ]

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        for sheet_name, df in sheets_to_export:
                            if not df.empty:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)

                                worksheet = writer.sheets[sheet_name]

                                max_row = len(df)
                                max_col = len(df.columns) - 1

                                worksheet.autofilter(0, 0, max_row, max_col)

                    output.seek(0)

                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"backup_inventory_{timestamp}.xlsx"

                    st.download_button(
                        label="⬇️ Klik untuk Download Excel",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.success("✅ File backup Excel siap didownload!")
                    if 'ready_to_download_excel' in st.session_state:
                        del st.session_state['ready_to_download_excel']

                except Exception as e:
                    st.error(f"❌ Error membuat backup Excel: {str(e)}. Coba jalankan perintah instalasi di atas!")

        with col2:
            if st.button("📄 Download Database File (.db)", use_container_width=True):
                st.session_state['ready_to_download_db'] = True

            if st.session_state.get('ready_to_download_db'):
                try:
                    with open('inventory_rumah.db', 'rb') as f:
                        db_data = f.read()

                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"inventory_database_{timestamp}.db"

                    st.download_button(
                        label="⬇️ Klik untuk Download DB",
                        data=db_data,
                        file_name=filename,
                        mime="application/x-sqlite3"
                    )

                    st.success("✅ Database file siap didownload!")
                    if 'ready_to_download_db' in st.session_state:
                        del st.session_state['ready_to_download_db']

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

            st.markdown("---")
            st.info("💡 **Tips:** Simpan backup secara rutin (minimal seminggu sekali) untuk keamanan data Anda.")

# Footer
st.markdown("""
<div style="text-align: center; color: #666; padding: 20px;">
    <h4>🏭 Aplikasi Inventory Gudang </h4>
    <p>📱 Kelola inventory Gudang Anda dengan mudah!</p>
    <br>
</div>
""", unsafe_allow_html=True)
