import pandas as pd
import streamlit as st
from datetime import datetime, time as dt_time, timedelta, date
import io
from collections import defaultdict
import random
import os
import warnings
from icalendar import Calendar, Event
import smtplib
from email.mime.text import MIMEText
import logging
from typing import Dict, List, Tuple, Optional, Any
try:
    from streamlit_calendar import calendar
except ImportError:
    calendar = None

# ========== SETUP LOGGING ==========
logging.basicConfig(
    filename='scheduler.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ========== KONFIGURASI UTAMA ==========
warnings.filterwarnings("ignore", category=UserWarning)

class Config:
    SEMESTER_KELAS = {"TI24": 1, "TI23": 3, "TI22": 5}
    KONSENTRASI_OPTIONS = ["AI", "software", "cybersecurity", "umum"]
    MAX_SKS_SEMESTER = 21
    MAX_SKS_DOSEN = 12
    DURASI_SKS = 50  # menit per SKS
    MAX_SCHEDULING_ATTEMPTS = 20
    PRIORITAS_RUANGAN_PREFIX = "B4"
    
    HARI_PRIORITAS = {
        'reguler': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'],
        'internasional': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'],
        'sabtu': ['Sabtu'],
        'karyawan': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'],
        'reguler malam': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
    }
    
    JAM_OPERASIONAL = {
        'reguler': (dt_time(8, 0), dt_time(17, 0)),
        'internasional': (dt_time(8, 0), dt_time(17, 0)),
        'sabtu': (dt_time(8, 0), dt_time(17, 0)),
        'karyawan': (dt_time(19, 0), dt_time(21, 0)),
        'reguler malam': (dt_time(19, 0), dt_time(21, 0))
    }
    
    WAKTU_TIDAK_BOLEH = {
        'Jumat': [(dt_time(11, 30), dt_time(13, 0))],  # Waktu sholat Jumat
    }
    
    ISTIRAHAT = [
        (dt_time(12, 0), dt_time(13, 0)),  # Istirahat siang
    ]
    
    WARNA_KELAS = {
        'Online': '#4e79a7',
        'Offline': '#f28e2b',
        'Karyawan': '#59a14f',
        'Sabtu': '#e15759',
        'Internasional': '#edc948',
        'AI': '#76b7b2',
        'software': '#59a14f',
        'cybersecurity': '#e15759',
        'Locked': '#d62728',
        'umum': '#bab0ac'
    }
    
    MATKUL_WAJIB = {
        1: ['Algoritma dan Struktur Data', 'Logika Informatika', 'Kalkulus', 'Statistika dan Probabilitas'],
        3: ['Metode Numerik', 'Pemrograman Berbasis Platform', 'Jaringan Komputer dan Keamanan Informasi', 'Rekayasa Perangkat Lunak'],
        5: ['Proyek Perangkat Lunak', 'Metodologi Penelitian', 'Basis Data']
    }

# ========== UTILITY FUNCTIONS ==========
def parse_time(time_str: str) -> dt_time:
    """Mengubah string waktu menjadi objek time dengan error handling"""
    if isinstance(time_str, str):
        try:
            if len(time_str.split(':')) == 2:
                return datetime.strptime(time_str, '%H:%M').time()
            return datetime.strptime(time_str, '%H:%M:%S').time()
        except ValueError as e:
            logging.warning(f"Gagal parse waktu: {time_str}, error: {str(e)}")
            return dt_time(8, 0)
    elif isinstance(time_str, dt_time):
        return time_str
    return dt_time(8, 0)

def generate_time_slots(
    jam_awal: dt_time, 
    jam_akhir: dt_time, 
    durasi_menit: int, 
    hari: str, 
    jenis_kelas: str
) -> List[Tuple[dt_time, dt_time]]:
    """Generate slot waktu yang tersedia dengan penyesuaian khusus"""
    slots = []
    jam_awal = parse_time(jam_awal)
    jam_akhir = parse_time(jam_akhir)
    
    dummy_date = date(2023, 1, 1)
    current_time = datetime.combine(dummy_date, jam_awal)
    end_time = datetime.combine(dummy_date, jam_akhir)
    durasi = timedelta(minutes=durasi_menit)
    
    # Penyesuaian khusus untuk jenis kelas
    if jenis_kelas.lower() in ['karyawan', 'reguler malam']:
        current_time = datetime.combine(dummy_date, dt_time(19, 0))
        end_time = datetime.combine(dummy_date, dt_time(21, 0))
    
    while current_time + durasi <= end_time:
        start = current_time.time()
        end = (current_time + durasi).time()
        
        # Cek bentrok dengan waktu istirahat
        is_istirahat = any(
            start <= istirahat_start < end or istirahat_start <= start < istirahat_end 
            for istirahat_start, istirahat_end in Config.ISTIRAHAT
        )
        
        if is_istirahat:
            current_time += timedelta(minutes=60)
            continue
        
        # Cek bentrok dengan waktu khusus
        if hari in Config.WAKTU_TIDAK_BOLEH:
            is_waktu_khusus = any(
                start <= waktu_start < end or waktu_start <= start < waktu_end
                for waktu_start, waktu_end in Config.WAKTU_TIDAK_BOLEH[hari]
            )
            if is_waktu_khusus:
                current_time += timedelta(minutes=90)
                continue
        
        slots.append((start, end))
        current_time += timedelta(minutes=durasi_menit + 10)  # Tambah jeda antar kelas
    
    return slots

def load_data() -> Tuple[pd.DataFrame, ...]:
    """Memuat data dari file Excel dengan validasi dan error handling"""
    try:
        file_path = "data.xlsx"
        if not os.path.exists(file_path):
            logging.error("File data.xlsx tidak ditemukan")
            return None, None, None, None, None, None, None
        
        # Baca semua sheet dengan validasi
        sheets = {
            "Kelas": None,
            "matakuliah": None,
            "Dosen": None,
            "dosen_matakuliah": None,
            "Hari": None,
            "ruangan": None,
            "availability": pd.DataFrame(columns=['dosen', 'hari', 'jam_mulai', 'jam_selesai'])
        }
        
        for sheet_name in sheets.keys():
            try:
                sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
                logging.info(f"Berhasil memuat sheet {sheet_name}")
            except Exception as e:
                logging.warning(f"Gagal memuat sheet {sheet_name}: {str(e)}")
                if sheet_name == "availability":
                    continue
                return None, None, None, None, None, None, None
        
        # Bersihkan data dosen_matakuliah
        if sheets["dosen_matakuliah"] is not None:
            sheets["dosen_matakuliah"] = sheets["dosen_matakuliah"].dropna()
        
        # Perbaiki typo di status matkul
        if sheets["matakuliah"] is not None and 'Status' in sheets["matakuliah"].columns:
            sheets["matakuliah"]['Status'] = sheets["matakuliah"]['Status'].str.replace('offlilne', 'offline')
        
        return tuple(sheets.values())
    except Exception as e:
        logging.error(f"Gagal memuat data: {str(e)}")
        return None, None, None, None, None, None, None

def save_to_excel(df: pd.DataFrame, sheet_name: str) -> bool:
    """Menyimpan dataframe ke sheet tertentu dalam file Excel dengan error handling"""
    try:
        file_path = "data.xlsx"
        
        # Baca file yang ada atau buat baru
        if os.path.exists(file_path):
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        logging.info(f"Berhasil menyimpan data ke sheet {sheet_name}")
        return True
    except Exception as e:
        logging.error(f"Gagal menyimpan data ke sheet {sheet_name}: {str(e)}")
        return False

class ResourceTracker:
    """Class untuk melacak penggunaan resource (kelas, dosen, ruangan)"""
    def __init__(self):
        self.kelas = defaultdict(list)
        self.dosen = defaultdict(list)
        self.ruangan = defaultdict(list)
    
    def add_schedule(
        self, 
        kelas: str, 
        dosen: str, 
        ruangan: str, 
        hari: str, 
        jam_mulai: dt_time, 
        jam_selesai: dt_time
    ) -> None:
        """Tambahkan jadwal ke resource tracker"""
        slot = {
            'hari': hari,
            'jam_mulai': jam_mulai,
            'jam_selesai': jam_selesai
        }
        
        self.kelas[kelas].append(slot)
        self.dosen[dosen].append(slot)
        if ruangan and ruangan != "Zoom":
            self.ruangan[ruangan].append(slot)
    
    def is_conflict(
        self, 
        kelas: str, 
        dosen: str, 
        ruangan: str, 
        hari: str, 
        jam_mulai: dt_time, 
        jam_selesai: dt_time
    ) -> bool:
        """Cek apakah ada konflik jadwal"""
        # Cek konflik kelas
        for slot in self.kelas.get(kelas, []):
            if slot['hari'] == hari:
                if (jam_mulai <= slot['jam_mulai'] < jam_selesai) or \
                   (jam_mulai < slot['jam_selesai'] <= jam_selesai):
                    return True
        
        # Cek konflik dosen
        for slot in self.dosen.get(dosen, []):
            if slot['hari'] == hari:
                if (jam_mulai <= slot['jam_mulai'] < jam_selesai) or \
                   (jam_mulai < slot['jam_selesai'] <= jam_selesai):
                    return True
        
        # Cek konflik ruangan (kecuali online)
        if ruangan and ruangan != "Zoom":
            for slot in self.ruangan.get(ruangan, []):
                if slot['hari'] == hari:
                    if (jam_mulai <= slot['jam_mulai'] < jam_selesai) or \
                       (jam_mulai < slot['jam_selesai'] <= jam_selesai):
                        return True
        
        return False

def is_dosen_busy(
    nama_dosen: str, 
    hari: str, 
    jam_mulai: dt_time, 
    jam_selesai: dt_time, 
    df_availability: pd.DataFrame, 
    resource_tracker: ResourceTracker
) -> bool:
    """Cek apakah dosen sibuk di waktu tertentu"""
    # Cek dari availability sheet
    if not df_availability.empty:
        busy = df_availability[
            (df_availability['dosen'] == nama_dosen) &
            (df_availability['hari'] == hari) &
            (
                ((df_availability['jam_mulai'] <= jam_mulai) & (df_availability['jam_selesai'] > jam_mulai)) |
                ((df_availability['jam_mulai'] < jam_selesai) & (df_availability['jam_selesai'] >= jam_selesai)) |
                ((jam_mulai <= df_availability['jam_mulai']) & (jam_selesai >= df_availability['jam_selesai']))
            )
        ]
        if not busy.empty:
            return True
    
    # Cek dari resource tracker
    for slot in resource_tracker.dosen.get(nama_dosen, []):
        if slot['hari'] == hari:
            if (jam_mulai <= slot['jam_mulai'] < jam_selesai) or \
               (jam_mulai < slot['jam_selesai'] <= jam_selesai) or \
               (slot['jam_mulai'] <= jam_mulai and slot['jam_selesai'] >= jam_selesai):
                return True
    return False

def cek_beban_dosen(nama_dosen: str, df_jadwal: pd.DataFrame) -> bool:
    """Cek beban mengajar dosen tidak melebihi MAX_SKS_DOSEN"""
    if df_jadwal.empty or 'Dosen' not in df_jadwal.columns:
        return True
    
    total_sks = df_jadwal[df_jadwal['Dosen'] == nama_dosen]['SKS'].sum()
    return total_sks < Config.MAX_SKS_DOSEN

def filter_matkul_by_konsentrasi(
    df_matkul: pd.DataFrame, 
    semester: int, 
    konsentrasi: str
) -> pd.DataFrame:
    """Filter matkul berdasarkan semester dan konsentrasi"""
    if df_matkul is None:
        return pd.DataFrame()
    
    # Filter berdasarkan semester
    filtered = df_matkul[df_matkul['semester'] == semester].copy()
    
    # Untuk semester 1 dan 3 hanya matkul umum
    if semester in [1, 3]:
        return filtered[filtered['Konsentrasi'] == 'umum']
    
    # Untuk semester 5 filter berdasarkan konsentrasi
    if konsentrasi != 'umum':
        filtered = filtered[
            filtered['Konsentrasi'].apply(
                lambda x: konsentrasi in [k.strip() for k in str(x).split(',')]
            )
        ]
    
    return filtered

def adjust_sks(df_matkul: pd.DataFrame) -> pd.DataFrame:
    """Sesuaikan SKS agar tidak melebihi batas maksimal"""
    if df_matkul.empty:
        return df_matkul
    
    total_sks = df_matkul['sks'].sum()
    if total_sks <= Config.MAX_SKS_SEMESTER:
        return df_matkul
    
    # Prioritaskan matkul wajib untuk dipertahankan
    wajib = []
    for matkul_list in Config.MATKUL_WAJIB.values():
        wajib.extend(matkul_list)
    
    df_wajib = df_matkul[df_matkul['nama'].str.contains('|'.join(wajib))]
    df_opsional = df_matkul[~df_matkul['nama'].str.contains('|'.join(wajib))]
    
    # Kurangi dari matkul opsional terlebih dahulu
    while total_sks > Config.MAX_SKS_SEMESTER and not df_opsional.empty:
        df_opsional = df_opsional.iloc[:-1]  # Hapus matkul terakhir
        total_sks = df_wajib['sks'].sum() + df_opsional['sks'].sum()
    
    return pd.concat([df_wajib, df_opsional])

def prioritize_matkul(df_matkul: pd.DataFrame) -> pd.DataFrame:
    """Prioritaskan matkul berdasarkan SKS dan status wajib"""
    df = df_matkul.copy()
    
    # Tambahkan kolom prioritas
    df['prioritas'] = df['sks'] * 2  # Matkul dengan SKS lebih besar lebih diprioritaskan
    
    # Tingkatkan prioritas untuk matkul wajib
    wajib = []
    for matkul_list in Config.MATKUL_WAJIB.values():
        wajib.extend(matkul_list)
    
    df['prioritas'] = df.apply(
        lambda x: x['prioritas'] + 10 if x['nama'] in wajib else x['prioritas'],
        axis=1
    )
    
    # Urutkan berdasarkan prioritas (descending)
    return df.sort_values('prioritas', ascending=False)

def validate_all_data(
    df_kelas: pd.DataFrame, 
    df_matkul: pd.DataFrame, 
    df_dosen: pd.DataFrame, 
    df_dosen_matkul: pd.DataFrame
) -> List[str]:
    """Validasi semua data sebelum generate jadwal"""
    errors = []
    
    if df_kelas is None or df_matkul is None or df_dosen is None or df_dosen_matkul is None:
        return ["Data tidak lengkap, pastikan semua sheet ada di file Excel"]
    
    # Validasi matkul wajib untuk semua semester
    for semester, matkul_wajib in Config.MATKUL_WAJIB.items():
        matkul_sem = df_matkul[df_matkul['semester'] == semester]
        for m in matkul_wajib:
            if m not in matkul_sem['nama'].values:
                errors.append(f"Matkul wajib {m} tidak ada di semester {semester}")
    
    # Validasi jenis kelas
    jenis_kelas_valid = ['reguler', 'internasional', 'sabtu', 'karyawan', 'reguler malam']
    invalid_jenis = df_kelas[~df_kelas['jenis'].isin(jenis_kelas_valid)]
    if not invalid_jenis.empty:
        errors.append(f"Jenis kelas tidak valid: {invalid_jenis['jenis'].unique()}")
    
    # Validasi hubungan dosen-matkul
    for _, matkul in df_matkul.iterrows():
        if df_dosen_matkul[df_dosen_matkul['id_matakuliah'] == matkul['id']].empty:
            errors.append(f"Matkul {matkul['nama']} tidak memiliki dosen")
    
    return errors

def schedule_matkul(
    matkul: pd.Series,
    kelas: pd.Series,
    df_dosen: pd.DataFrame,
    df_dosen_matkul: pd.DataFrame,
    df_ruangan: pd.DataFrame,
    df_availability: pd.DataFrame,
    resource_tracker: ResourceTracker,
    ruangan_prioritas: List[str]
) -> Dict[str, Any]:
    """Coba menjadwalkan satu mata kuliah"""
    nama_kelas = kelas['nama']
    jenis_kelas = kelas['jenis']
    konsentrasi = kelas.get('konsentrasi', 'umum')
    
    # Tentukan apakah harus offline atau bisa online
    must_offline = any(x in matkul['nama'].lower() for x in ['praktikum', 'lab', 'jaringan'])
    
    if must_offline:
        is_online = False
        ruangan_options = ruangan_prioritas.copy()
        random.shuffle(ruangan_options)
    else:
        is_online = matkul['Status'].lower().strip() == 'online'
        ruangan_options = ["Zoom"] if is_online else ruangan_prioritas.copy()
        if not is_online:
            random.shuffle(ruangan_options)

    # Atur hari tersedia
    hari_tersedia = Config.HARI_PRIORITAS.get(jenis_kelas.lower(), ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'])
    
    # Daftar dosen yang tersedia untuk matkul ini
    dosen_ids = df_dosen_matkul[df_dosen_matkul['id_matakuliah'] == matkul['id']]['id_dosen']
    dosen_tersedia = df_dosen[df_dosen['id'].isin(dosen_ids)]
    
    if dosen_tersedia.empty:
        return {
            'Kelas': nama_kelas,
            'Konsentrasi': konsentrasi,
            'Hari': 'Cek EdLink',
            'Jam': 'Cek EdLink',
            'Mata Kuliah': matkul['nama'],
            'Dosen': 'Belum Ditentukan',
            'Ruangan': 'Zoom' if is_online else 'Cek EdLink',
            'SKS': matkul['sks'],
            'Semester': matkul['semester'],
            'Status': 'Online' if is_online else 'Offline',
            'Keterangan': '⚠️ Tanpa Dosen',
            'Warna': Config.WARNA_KELAS['Online'] if is_online else Config.WARNA_KELAS.get(konsentrasi, Config.WARNA_KELAS['Offline']),
            'is_locked': False
        }
    
    # Coba menjadwalkan
    for _ in range(Config.MAX_SCHEDULING_ATTEMPTS):
        random.shuffle(hari_tersedia)
        dosen_tersedia = dosen_tersedia.sample(frac=1)  # Acak urutan dosen
        
        for hari in hari_tersedia:
            jam_awal, jam_akhir = Config.JAM_OPERASIONAL.get(jenis_kelas.lower(), (dt_time(8, 0), dt_time(17, 0)))
            
            possible_slots = generate_time_slots(jam_awal, jam_akhir, matkul['sks'] * Config.DURASI_SKS, hari, jenis_kelas)
            
            for jam_mulai, jam_selesai in possible_slots:
                for _, dosen in dosen_tersedia.iterrows():
                    nama_dosen = dosen['nama']
                    
                    # Cek ketersediaan dosen
                    if is_dosen_busy(nama_dosen, hari, jam_mulai, jam_selesai, df_availability, resource_tracker):
                        continue
                        
                    if is_online:
                        if not resource_tracker.is_conflict(nama_kelas, nama_dosen, None, hari, jam_mulai, jam_selesai):
                            resource_tracker.add_schedule(nama_kelas, nama_dosen, "Zoom", hari, jam_mulai, jam_selesai)
                            
                            return {
                                'Kelas': nama_kelas,
                                'Konsentrasi': konsentrasi,
                                'Hari': hari,
                                'Jam': f"{jam_mulai.strftime('%H:%M')}-{jam_selesai.strftime('%H:%M')}",
                                'Mata Kuliah': matkul['nama'],
                                'Dosen': nama_dosen,
                                'Ruangan': "Zoom",
                                'SKS': matkul['sks'],
                                'Semester': matkul['semester'],
                                'Status': 'Online',
                                'Keterangan': '✅',
                                'Warna': Config.WARNA_KELAS['Online'],
                                'is_locked': False
                            }
                    else:
                        for ruangan in ruangan_options:
                            if not resource_tracker.is_conflict(nama_kelas, nama_dosen, ruangan, hari, jam_mulai, jam_selesai):
                                resource_tracker.add_schedule(nama_kelas, nama_dosen, ruangan, hari, jam_mulai, jam_selesai)
                                
                                return {
                                    'Kelas': nama_kelas,
                                    'Konsentrasi': konsentrasi,
                                    'Hari': hari,
                                    'Jam': f"{jam_mulai.strftime('%H:%M')}-{jam_selesai.strftime('%H:%M')}",
                                    'Mata Kuliah': matkul['nama'],
                                    'Dosen': nama_dosen,
                                    'Ruangan': ruangan,
                                    'SKS': matkul['sks'],
                                    'Semester': matkul['semester'],
                                    'Status': 'Offline',
                                    'Keterangan': '✅',
                                    'Warna': Config.WARNA_KELAS.get(konsentrasi, Config.WARNA_KELAS['Offline']),
                                    'is_locked': False
                                }
    
    # Jika gagal setelah semua percobaan
    return {
        'Kelas': nama_kelas,
        'Konsentrasi': konsentrasi,
        'Hari': 'Cek EdLink',
        'Jam': 'Cek EdLink',
        'Mata Kuliah': matkul['nama'],
        'Dosen': random.choice(dosen_tersedia['nama'].tolist()),
        'Ruangan': 'Zoom' if is_online else 'Cek EdLink',
        'SKS': matkul['sks'],
        'Semester': matkul['semester'],
        'Status': 'Online' if is_online else 'Offline',
        'Keterangan': f'⚠️ Gagal setelah {Config.MAX_SCHEDULING_ATTEMPTS}x attempt',
        'Warna': Config.WARNA_KELAS['Online'] if is_online else Config.WARNA_KELAS.get(konsentrasi, Config.WARNA_KELAS['Offline']),
        'is_locked': False
    }

def generate_jadwal() -> Optional[pd.DataFrame]:
    """Generate jadwal kuliah secara otomatis dengan penjadwalan yang lebih cerdas"""
    # Load data
    df_kelas, df_matkul, df_dosen, df_dosen_matkul, df_hari, df_ruangan, df_availability = load_data()
    if df_kelas is None or df_matkul is None or df_dosen is None or df_dosen_matkul is None:
        st.error("Data tidak lengkap, pastikan semua sheet ada di file Excel")
        return None

    # Validasi data
    validation_errors = validate_all_data(df_kelas, df_matkul, df_dosen, df_dosen_matkul)
    if validation_errors:
        st.error("Error validasi data:\n- " + "\n- ".join(validation_errors))
        return None

    # Inisialisasi konsentrasi jika belum ada
    if 'konsentrasi' not in df_kelas.columns:
        df_kelas['konsentrasi'] = df_kelas['nama'].apply(
            lambda x: 'umum' if x[:4] in ['TI24', 'TI23'] else random.choice(Config.KONSENTRASI_OPTIONS)
        )
    
    # Prioritaskan ruangan
    ruangan_prioritas = [r for r in df_ruangan['nama'] if Config.PRIORITAS_RUANGAN_PREFIX in r]
    ruangan_lain = [r for r in df_ruangan['nama'] if Config.PRIORITAS_RUANGAN_PREFIX not in r]
    ruangan_prioritas += ruangan_lain
    
    jadwal_all = []
    resource_tracker = ResourceTracker()
    
    progress_bar = st.progress(0)
    total_kelas = len(df_kelas)
    
    for i, (_, kelas) in enumerate(df_kelas.iterrows()):
        nama_kelas = kelas['nama']
        jenis_kelas = kelas['jenis']
        konsentrasi = kelas['konsentrasi']
        prefix_kelas = nama_kelas[:4]
        
        if prefix_kelas not in Config.SEMESTER_KELAS:
            continue

        semester = Config.SEMESTER_KELAS[prefix_kelas]
        
        # Filter matkul berdasarkan semester dan konsentrasi
        matkul_kelas = filter_matkul_by_konsentrasi(df_matkul, semester, konsentrasi)
        
        # Sesuaikan SKS
        matkul_kelas = adjust_sks(matkul_kelas)
        
        if matkul_kelas.empty:
            st.warning(f"Tidak ada mata kuliah untuk semester {semester}")
            continue
        
        # Prioritaskan matkul (yang SKS besar dan wajib dijadwal lebih awal)
        matkul_kelas = prioritize_matkul(matkul_kelas)
        
        jadwal_kelas = []
        
        for _, matkul in matkul_kelas.iterrows():
            jadwal = schedule_matkul(
                matkul, kelas, df_dosen, df_dosen_matkul, 
                df_ruangan, df_availability, resource_tracker, 
                ruangan_prioritas
            )
            jadwal_kelas.append(jadwal)
        
        jadwal_all.extend(jadwal_kelas)
        progress_bar.progress((i + 1) / total_kelas)
    
    return pd.DataFrame(jadwal_all)

def jadwal_to_calendar_events(jadwal_df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Konversi jadwal ke format event kalender"""
    events = []
    
    if jadwal_df.empty:
        return events
    
    for _, row in jadwal_df.iterrows():
        if row['Hari'] == 'Cek EdLink':
            continue
            
        hari_to_num = {'Senin': 0, 'Selasa': 1, 'Rabu': 2, 'Kamis': 3, 
                      'Jumat': 4, 'Sabtu': 5, 'Minggu': 6}
        
        try:
            jam_parts = row['Jam'].split('-')
            if len(jam_parts) != 2:
                continue
                
            start_time = datetime.strptime(jam_parts[0], '%H:%M').time()
            end_time = datetime.strptime(jam_parts[1], '%H:%M').time()
            
            day_number = 2 + hari_to_num.get(row['Hari'], 0)
            events.append({
                'title': f"{row['Mata Kuliah']} ({row['Kelas']})",
                'start': f"2023-01-{day_number:02d}T{start_time.strftime('%H:%M:%S')}",
                'end': f"2023-01-{day_number:02d}T{end_time.strftime('%H:%M:%S')}",
                'color': row.get('Warna', Config.WARNA_KELAS['Offline']),
                'extendedProps': {
                    'dosen': row.get('Dosen', ''),
                    'sks': row.get('SKS', 0),
                    'status': row.get('Status', 'Offline'),
                    'ruangan': row.get('Ruangan', ''),
                    'keterangan': row.get('Keterangan', ''),
                    'konsentrasi': row.get('Konsentrasi', 'umum')
                }
            })
        except Exception as e:
            logging.warning(f"Gagal memproses jadwal: {row.get('Mata Kuliah', 'Unknown')}. Error: {str(e)}")
    
    return events

def show_calendar_view(jadwal_df: pd.DataFrame) -> None:
    """Tampilkan jadwal dalam bentuk kalender interaktif"""
    if calendar is None:
        st.error("Fitur kalender membutuhkan package streamlit-calendar. Install dengan: pip install streamlit-calendar")
        return
    
    if jadwal_df.empty:
        st.warning("Tidak ada jadwal untuk ditampilkan")
        return
    
    st.subheader("📅 Kalender Interaktif")
    
    tab1, tab2 = st.tabs(["Mingguan", "Bulanan"])
    
    with tab1:
        calendar_options = {
            "editable": False,
            "selectable": True,
            "headerToolbar": {
                "left": "today prev,next",
                "center": "title",
                "right": "timeGridWeek,dayGridMonth"
            },
            "initialView": "timeGridWeek",
            "slotMinTime": "07:00:00",
            "slotMaxTime": "22:00:00",
            "eventClick": """
                function(info) {
                    alert(
                        'Detail Jadwal:\\n\\n' +
                        'Mata Kuliah: ' + info.event.title + '\\n' +
                        'Dosen: ' + info.event.extendedProps.dosen + '\\n' +
                        'SKS: ' + info.event.extendedProps.sks + '\\n' +
                        'Status: ' + info.event.extendedProps.status + '\\n' +
                        'Ruangan: ' + info.event.extendedProps.ruangan + '\\n' +
                        'Konsentrasi: ' + (info.event.extendedProps.konsentrasi || 'umum') + '\\n' +
                        'Keterangan: ' + info.event.extendedProps.keterangan
                    );
                }
            """
        }
        calendar(events=jadwal_to_calendar_events(jadwal_df), 
                options=calendar_options, 
                key="week_calendar")

    with tab2:
        calendar_options = {
            "initialView": "dayGridMonth",
            "headerToolbar": {
                "left": "today prev,next",
                "center": "title",
                "right": "dayGridMonth,timeGridWeek"
            }
        }
        calendar(events=jadwal_to_calendar_events(jadwal_df), 
                options=calendar_options, 
                key="month_calendar")

def edit_jadwal_manual(
    jadwal_df: pd.DataFrame, 
    df_dosen: pd.DataFrame, 
    df_ruangan: pd.DataFrame
) -> pd.DataFrame:
    """Fitur edit jadwal manual"""
    if jadwal_df.empty:
        return jadwal_df
        
    st.subheader("✏️ Edit Jadwal Manual")
    
    edited_df = st.data_editor(
        jadwal_df,
        disabled=["Kelas", "Mata Kuliah", "SKS", "Semester", "Konsentrasi"],
        column_config={
            "Hari": st.column_config.SelectboxColumn(
                options=["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"]
            ),
            "Jam": st.column_config.TextColumn(
                help="Format: HH:MM-HH:MM"
            ),
            "Dosen": st.column_config.SelectboxColumn(
                options=df_dosen['nama'].tolist() if df_dosen is not None else []
            ),
            "Ruangan": st.column_config.SelectboxColumn(
                options=["Zoom"] + df_ruangan['nama'].tolist() if df_ruangan is not None else ["Zoom"]
            ),
            "is_locked": st.column_config.CheckboxColumn(
                label="Lock",
                help="Centang untuk mengunci jadwal ini"
            ),
            "Warna": st.column_config.Column(disabled=True)
        },
        hide_index=True,
        use_container_width=True
    )
    
    if st.button("💾 Simpan Perubahan Manual"):
        return edited_df
    return jadwal_df

def generate_report(jadwal_df: pd.DataFrame) -> Dict[str, Any]:
    """Generate laporan analisis jadwal"""
    if jadwal_df.empty:
        return {
            "total_kelas": 0,
            "total_matkul": 0,
            "total_dosen": 0,
            "konflik_jadwal": [],
            "beban_dosen": {},
            "penggunaan_ruangan": {}
        }
    
    report = {
        "total_kelas": jadwal_df['Kelas'].nunique(),
        "total_matkul": jadwal_df['Mata Kuliah'].nunique(),
        "total_dosen": jadwal_df['Dosen'].nunique(),
        "konflik_jadwal": [],
        "beban_dosen": jadwal_df.groupby('Dosen')['SKS'].sum().sort_values(ascending=False).to_dict(),
        "penggunaan_ruangan": jadwal_df[jadwal_df['Ruangan'] != 'Zoom']['Ruangan'].value_counts().to_dict()
    }
    
    return report

def send_notification(email: str, subject: str, message: str) -> bool:
    """Kirim notifikasi via email"""
    try:
        msg = MIMEText(message)
        msg['Subject'] = subject
        msg['From'] = "sistem_penjadwalan@univ.ac.id"
        msg['To'] = email
        
        with smtplib.SMTP('smtp.example.com', 587) as server:
            server.starttls()
            server.login("username", "password")
            server.send_message(msg)
        return True
    except Exception as e:
        logging.error(f"Gagal mengirim email: {str(e)}")
        return False

def export_to_ical(jadwal_df: pd.DataFrame) -> bytes:
    """Ekspor jadwal ke format iCal"""
    cal = Calendar()
    cal.add('prodid', '-//Jadwal Kuliah//univ.ac.id//')
    cal.add('version', '2.0')
    
    if jadwal_df.empty:
        return cal.to_ical()
    
    for _, row in jadwal_df.iterrows():
        if row['Hari'] == 'Cek EdLink':
            continue
            
        try:
            event = Event()
            event.add('summary', f"{row['Mata Kuliah']} ({row['Kelas']})")
            event.add('description', f"Dosen: {row['Dosen']}\nRuangan: {row['Ruangan']}")
            
            hari_to_num = {'Senin': 0, 'Selasa': 1, 'Rabu': 2, 
                          'Kamis': 3, 'Jumat': 4, 'Sabtu': 5}
            start_date = date(2023, 8, 1)  # Contoh: 1 Agustus 2023
            days_to_add = hari_to_num.get(row['Hari'], 0)
            event_date = start_date + timedelta(days=days_to_add)
            
            jam_parts = row['Jam'].split('-')
            if len(jam_parts) != 2:
                continue
                
            start_time = datetime.strptime(jam_parts[0], '%H:%M').time()
            end_time = datetime.strptime(jam_parts[1], '%H:%M').time()
            
            event.add('dtstart', datetime.combine(event_date, start_time))
            event.add('dtend', datetime.combine(event_date, end_time))
            event.add('location', row['Ruangan'])
            
            cal.add_component(event)
        except Exception as e:
            logging.warning(f"Gagal memproses jadwal {row['Mata Kuliah']}: {str(e)}")
    
    return cal.to_ical()

def main():
    st.set_page_config(layout="wide", page_title="Sistem Penjadwalan Kuliah TI", page_icon="🎓")

    # Initialize session state variables
    if 'jadwal_df' not in st.session_state:
        st.session_state.jadwal_df = None

    # CSS untuk tampilan lebih baik
    st.markdown("""
    <style>
        .stRadio [role=radiogroup] {
            gap: 0.5rem;
        }
        .stRadio [role=radio] {
            padding: 0.5rem;
            border-radius: 0.5rem;
            border: 1px solid #e0e0e0;
            transition: all 0.2s;
        }
        .stRadio [role=radio]:hover {
            background: #f5f5f5;
        }
        .stRadio [role=radio][aria-checked=true] {
            background: #f0f8ff;
            border-color: #1e90ff;
            color: #1e90ff;
        }
        .fc-event {
            cursor: pointer;
            font-size: 0.85em;
            padding: 2px;
        }
        .stDataFrame {
            font-size: 0.9em;
        }
        .locked-row { 
            background-color: #ffcccc; 
        }
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.image("https://via.placeholder.com/150x50.png?text=UNIVERSITAS", width=150)
        st.title("Menu")
        
        menu_option = st.radio(
            "Pilihan Menu",
            ["🏠 Beranda", "📅 Generate Jadwal", "👨‍🏫 Manajemen Dosen", "⚙️ Kelola Konsentrasi", "🗓️ Kalender Interaktif", "✏️ Edit Manual", "📊 Laporan"],
            label_visibility="collapsed"
        )
        
        st.divider()
        st.caption("Jurusan Teknik Informatika")
        st.caption(f"Versi {datetime.now().strftime('%Y-%m-%d')}")

    if menu_option == "🏠 Beranda":
        st.title("🎓 Sistem Penjadwalan Kuliah TI")
        st.write("""
        Selamat datang di Sistem Penjadwalan Kuliah Jurusan Teknik Informatika.
        Gunakan menu di sidebar untuk mengakses fitur-fitur berikut:
        
        - **Generate Jadwal**: Membuat jadwal kuliah otomatis dengan penyesuaian khusus untuk setiap jenis kelas
        - **Manajemen Dosen**: Mengatur data dosen dan matakuliah
        - **Kelola Konsentrasi**: Mengatur konsentrasi untuk setiap kelas
        - **Kalender Interaktif**: Lihat jadwal dalam tampilan kalender yang informatif
        - **Edit Manual**: Edit jadwal secara manual
        - **Laporan**: Analisis jadwal dan beban mengajar
        """)
        
        st.info("""
        **Fitur Unggulan:**
        1. Penjadwalan cerdas untuk berbagai jenis kelas (Reguler, Sabtu, Karyawan, Internasional)
        2. Pembagian matkul berdasarkan konsentrasi (AI, Software, Cybersecurity)
        3. Pembatasan SKS maksimal per semester (maks 21 SKS)
        4. Fleksibilitas mode online/offline
        5. Prioritas ruangan dan dosen
        6. Kalender interaktif dengan detail lengkap
        7. Ekspor ke format iCal
        8. Notifikasi email
        """)

    elif menu_option == "📅 Generate Jadwal":
        st.title("📅 Generate Jadwal Kuliah")
        
        df_kelas, _, _, df_dosen_matkul, _, _, _ = load_data()
        errors = []
        if df_kelas is None:
            errors.append("❌ File data.xlsx tidak ditemukan atau format tidak sesuai!")
        if df_dosen_matkul is None or df_dosen_matkul.empty:
            errors.append("❌ Tidak ada hubungan dosen-matkul!")
        
        if errors:
            st.error("\n".join(errors))
        else:
            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button("🔄 Generate Jadwal Baru", type="primary", use_container_width=True):
                    with st.spinner("Membuat jadwal..."):
                        st.session_state.jadwal_df = generate_jadwal()
                        if st.session_state.jadwal_df is not None:
                            st.toast("Jadwal berhasil dibuat!", icon="✅")
            
            with col2:
                if st.button("🔄 Reset Jadwal", type="secondary", use_container_width=True):
                    st.session_state.jadwal_df = None
                    st.rerun()
        
        if 'jadwal_df' in st.session_state and st.session_state.jadwal_df is not None:
            st.success("Jadwal berhasil dibuat!")
            
            with st.expander("🔍 Filter Jadwal", expanded=True):
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    filter_kelas = st.multiselect(
                        "Kelas",
                        options=st.session_state.jadwal_df['Kelas'].unique()
                    )
                with col2:
                    filter_konsentrasi = st.multiselect(
                        "Konsentrasi",
                        options=st.session_state.jadwal_df['Konsentrasi'].unique()
                    )
                with col3:
                    filter_semester = st.multiselect(
                        "Semester",
                        options=st.session_state.jadwal_df['Semester'].unique()
                    )
                with col4:
                    filter_status = st.multiselect(
                        "Status",
                        options=st.session_state.jadwal_df['Status'].unique(),
                        default=['Online', 'Offline']
                    )
                with col5:
                    filter_hari = st.multiselect(
                        "Hari",
                        options=st.session_state.jadwal_df['Hari'].unique()
                    )
            
            filtered_df = st.session_state.jadwal_df.copy()
            if filter_kelas:
                filtered_df = filtered_df[filtered_df['Kelas'].isin(filter_kelas)]
            if filter_konsentrasi:
                filtered_df = filtered_df[filtered_df['Konsentrasi'].isin(filter_konsentrasi)]
            if filter_semester:
                filtered_df = filtered_df[filtered_df['Semester'].isin(filter_semester)]
            if filter_status:
                filtered_df = filtered_df[filtered_df['Status'].isin(filter_status)]
            if filter_hari:
                filtered_df = filtered_df[filtered_df['Hari'].isin(filter_hari)]
            
            st.dataframe(
                filtered_df.sort_values(['Kelas', 'Hari', 'Jam']),
                height=600,
                use_container_width=True,
                column_config={
                    "Keterangan": st.column_config.Column(width="medium"),
                    "Warna": st.column_config.Column(disabled=True)
                }
            )
            
            # Ekspor ke iCal
            st.download_button(
                label="📅 Ekspor ke Kalender (iCal)",
                data=export_to_ical(filtered_df),
                file_name="jadwal_kuliah.ics",
                mime="text/calendar"
            )
            
            excel_buffer = io.BytesIO()
            filtered_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            
            st.download_button(
                label="💾 Download Jadwal (Excel)",
                data=excel_buffer,
                file_name="jadwal_kuliah.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    elif menu_option == "👨‍🏫 Manajemen Dosen":
        st.title("👨‍🏫 Manajemen Dosen & Matakuliah")
        
        tab1, tab2, tab3 = st.tabs(["Tambah Dosen", "Tambah Matakuliah", "Hubungkan Dosen-Matkul"])
        
        df_kelas, df_matkul, df_dosen, df_dosen_matkul, df_hari, df_ruangan, df_availability = load_data()
        
        with tab1:
            st.subheader("Tambah Dosen Baru")
            with st.form("form_dosen"):
                nama = st.text_input("Nama Dosen*")
                submitted = st.form_submit_button("Simpan")
                
                if submitted:
                    if not nama:
                        st.error("Nama wajib diisi!")
                    elif df_dosen is not None and not df_dosen.empty and nama in df_dosen['nama'].values:
                        st.warning("⚠️ Dosen sudah terdaftar!")
                    else:
                        new_id = max(df_dosen['id']) + 1 if df_dosen is not None and not df_dosen.empty else 1
                        new_dosen = pd.DataFrame([{'id': new_id, 'nama': nama}])
                        updated_dosen = pd.concat([df_dosen, new_dosen], ignore_index=True) if df_dosen is not None else new_dosen
                        if save_to_excel(updated_dosen, "Dosen"):
                            st.success("Dosen berhasil ditambahkan!")
                            st.rerun()
            
            st.divider()
            st.subheader("Daftar Dosen")
            if df_dosen is not None:
                st.dataframe(df_dosen, use_container_width=True, hide_index=True)
                
                if not df_dosen.empty:
                    selected_dosen = st.selectbox("Pilih Dosen untuk Dihapus", df_dosen['nama'])
                    if st.button("🗑️ Hapus Dosen", type="secondary"):
                        updated_dosen = df_dosen[df_dosen['nama'] != selected_dosen]
                        if save_to_excel(updated_dosen, "Dosen"):
                            st.success("Dosen dihapus!")
                            st.rerun()
            else:
                st.warning("Data dosen tidak tersedia")
        
        with tab2:
            st.subheader("Tambah Matakuliah Baru")
            with st.form("form_matkul"):
                nama = st.text_input("Nama Matakuliah*")
                sks = st.number_input("SKS*", min_value=1, max_value=6, step=1)
                semester = st.number_input("Semester*", min_value=1, max_value=8, step=1)
                status = st.selectbox("Status*", ["Online", "Offline"])
                konsentrasi = st.text_input("Konsentrasi (pisahkan dengan koma jika multiple)", "umum")
                submitted = st.form_submit_button("Simpan")
                
                if submitted:
                    if not nama:
                        st.error("Nama matkul wajib diisi!")
                    elif df_matkul is not None and not df_matkul.empty and nama in df_matkul['nama'].values:
                        st.warning("⚠️ Matakuliah sudah terdaftar!")
                    else:
                        new_id = max(df_matkul['id']) + 1 if df_matkul is not None and not df_matkul.empty else 1
                        new_matkul = pd.DataFrame([{
                            'id': new_id, 'nama': nama, 'sks': sks, 
                            'semester': semester, 'Status': status,
                            'Konsentrasi': konsentrasi
                        }])
                        updated_matkul = pd.concat([df_matkul, new_matkul], ignore_index=True) if df_matkul is not None else new_matkul
                        if save_to_excel(updated_matkul, "matakuliah"):
                            st.success("Matakuliah berhasil ditambahkan!")
                            st.rerun()
            
            st.divider()
            st.subheader("Daftar Matakuliah")
            if df_matkul is not None:
                st.dataframe(df_matkul, use_container_width=True, hide_index=True)
            else:
                st.warning("Data matakuliah tidak tersedia")
        
        with tab3:
            st.subheader("Hubungkan Dosen dengan Matakuliah")
            col1, col2 = st.columns(2)
            with col1:
                selected_dosen = st.selectbox("Pilih Dosen", df_dosen['nama'] if df_dosen is not None else [])
                dosen_id = df_dosen[df_dosen['nama'] == selected_dosen]['id'].iloc[0] if df_dosen is not None and not df_dosen.empty and selected_dosen else None
            with col2:
                selected_matkul = st.selectbox("Pilih Matakuliah", df_matkul['nama'] if df_matkul is not None else [])
                matkul_id = df_matkul[df_matkul['nama'] == selected_matkul]['id'].iloc[0] if df_matkul is not None and not df_matkul.empty and selected_matkul else None
            
            if st.button("🔗 Hubungkan"):
                if dosen_id is None or matkul_id is None:
                    st.error("Pilih dosen dan matakuliah yang valid!")
                else:
                    if df_dosen_matkul is not None and not df_dosen_matkul.empty and ((df_dosen_matkul['id_dosen'] == dosen_id) & (df_dosen_matkul['id_matakuliah'] == matkul_id)).any():
                        st.warning("⚠️ Hubungan sudah ada!")
                    else:
                        new_id = max(df_dosen_matkul['id']) + 1 if df_dosen_matkul is not None and not df_dosen_matkul.empty else 1
                        new_link = pd.DataFrame([{'id': new_id, 'id_dosen': dosen_id, 'id_matakuliah': matkul_id}])
                        updated_link = pd.concat([df_dosen_matkul, new_link], ignore_index=True) if df_dosen_matkul is not None else new_link
                        if save_to_excel(updated_link, "dosen_matakuliah"):
                            st.success("Berhasil dihubungkan!")
                            st.rerun()
            
            st.divider()
            st.subheader("Daftar Pengajaran")
            if df_dosen_matkul is not None and not df_dosen_matkul.empty and df_dosen is not None and df_matkul is not None:
                df_tampil = df_dosen_matkul.merge(
                    df_dosen, left_on='id_dosen', right_on='id'
                ).merge(
                    df_matkul, left_on='id_matakuliah', right_on='id'
                )[['nama_x', 'nama_y', 'sks', 'semester', 'Konsentrasi']].rename(
                    columns={'nama_x': 'Dosen', 'nama_y': 'Matakuliah'}
                )
                st.dataframe(df_tampil, use_container_width=True, hide_index=True)
            else:
                st.info("Belum ada hubungan dosen-matakuliah.")

    elif menu_option == "⚙️ Kelola Konsentrasi":
        st.title("⚙️ Kelola Konsentrasi Kelas")
        
        df_kelas = load_data()[0]
        if df_kelas is None:
            st.error("Data kelas tidak dapat dimuat")
            return
        
        # Inisialisasi konsentrasi jika belum ada
        if 'konsentrasi' not in df_kelas.columns:
            df_kelas['konsentrasi'] = df_kelas['nama'].apply(
                lambda x: 'umum' if x[:4] in ['TI24', 'TI23'] else random.choice(['AI', 'software', 'cybersecurity'])
            )
        
        st.subheader("Edit Konsentrasi Kelas")
        edited_df = st.data_editor(
            df_kelas[['nama', 'jenis', 'konsentrasi']],
            column_config={
                "konsentrasi": st.column_config.SelectboxColumn(
                    "Konsentrasi",
                    options=Config.KONSENTRASI_OPTIONS,
                    required=True
                )
            },
            hide_index=True,
            use_container_width=True
        )
        
        if st.button("💾 Simpan Perubahan Konsentrasi"):
            df_kelas['konsentrasi'] = edited_df['konsentrasi']
            if save_to_excel(df_kelas, "Kelas"):
                st.success("Konsentrasi berhasil disimpan!")
                st.rerun()
        
        st.divider()
        st.subheader("Distribusi Konsentrasi")
        konsentrasi_dist = edited_df['konsentrasi'].value_counts()
        st.bar_chart(konsentrasi_dist)

    elif menu_option == "🗓️ Kalender Interaktif":
        st.title("🗓️ Kalender Interaktif")
        
        if 'jadwal_df' not in st.session_state or st.session_state.jadwal_df is None:
            st.warning("Generate jadwal terlebih dahulu")
        else:
            show_calendar_view(st.session_state.jadwal_df)

    elif menu_option == "✏️ Edit Manual":
        st.title("✏️ Edit Jadwal Manual")
        
        if 'jadwal_df' not in st.session_state or st.session_state.jadwal_df is None:
            st.warning("Generate jadwal terlebih dahulu")
        else:
            df_dosen = load_data()[2]
            df_ruangan = load_data()[5]
            st.session_state.jadwal_df = edit_jadwal_manual(
                st.session_state.jadwal_df, 
                df_dosen, 
                df_ruangan
            )

    elif menu_option == "📊 Laporan":
        st.title("📊 Laporan dan Analisis")
        
        if 'jadwal_df' not in st.session_state or st.session_state.jadwal_df is None:
            st.warning("Belum ada jadwal yang digenerate")
        else:
            report = generate_report(st.session_state.jadwal_df)
            
            st.subheader("Statistik Dosen")
            if report['beban_dosen']:
                st.bar_chart(pd.DataFrame.from_dict(report['beban_dosen'], orient='index'))
            else:
                st.warning("Tidak ada data beban dosen")
            
            st.subheader("Penggunaan Ruangan")
            if report['penggunaan_ruangan']:
                st.bar_chart(pd.DataFrame.from_dict(report['penggunaan_ruangan'], orient='index'))
            else:
                st.warning("Tidak ada data penggunaan ruangan")
            
            if st.button("📧 Kirim Laporan ke Email"):
                if send_notification("admin@univ.ac.id", "Laporan Jadwal Kuliah", str(report)):
                    st.success("Laporan terkirim!")

    if st.sidebar.checkbox("🔍 Tampilkan Data Mentah"):
        st.title("Data Mentah")
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Kelas", "Mata Kuliah", "Dosen", "Ruangan", "Ketersediaan"])
        with tab1:
            df_kelas = load_data()[0]
            if df_kelas is not None:
                st.dataframe(df_kelas, use_container_width=True)
            else:
                st.warning("Data kelas tidak tersedia")
        with tab2:
            df_matkul = load_data()[1]
            if df_matkul is not None:
                st.dataframe(df_matkul, use_container_width=True)
            else:
                st.warning("Data matakuliah tidak tersedia")
        with tab3:
            df_dosen = load_data()[2]
            if df_dosen is not None:
                st.dataframe(df_dosen, use_container_width=True)
            else:
                st.warning("Data dosen tidak tersedia")
        with tab4:
            df_ruangan = load_data()[5]
            if df_ruangan is not None:
                st.dataframe(df_ruangan, use_container_width=True)
            else:
                st.warning("Data ruangan tidak tersedia")
        with tab5:
            df_availability = load_data()[6]
            if df_availability is not None and not df_availability.empty:
                st.dataframe(df_availability, use_container_width=True)
            else:
                st.info("Tidak ada data ketersediaan")

if __name__ == "__main__":
    main()