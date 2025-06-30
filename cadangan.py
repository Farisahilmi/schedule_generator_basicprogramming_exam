import pandas as pd
import streamlit as st
from datetime import datetime, time as dt_time, timedelta
import io
from collections import defaultdict
import random
import os
import json
import warnings
try:
    from streamlit_calendar import calendar
except ImportError:
    calendar = None

# ========== KONFIGURASI UTAMA ==========
warnings.filterwarnings("ignore", category=UserWarning)

SEMESTER_KELAS = {"TI24": 1, "TI23": 3, "TI22": 5}
KONSENTRASI_OPTIONS = ["AI", "software", "cybersecurity"]
MAX_SKS_SEMESTER = 21

HARI_PRIORITAS = {
    'reguler': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'],
    'internasional': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'],
    'sabtu': ['Sabtu'],
    'karyawan': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'],
    'reguler malam': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
}

JAM_OPERASIONAL = {
    'reguler': (dt_time(8, 0), dt_time(21, 0)),
    'internasional': (dt_time(8, 0), dt_time(21, 0)),
    'sabtu': (dt_time(8, 0), dt_time(21, 0)),
    'karyawan': (dt_time(8, 0), dt_time(21, 0)),
    'reguler malam': (dt_time(19, 0), dt_time(21, 0)),
    'reguler_D': (dt_time(8, 0), dt_time(20, 0))  # Khusus kelas D
}

WAKTU_TIDAK_BOLEH = {
    'Jumat': [(dt_time(11, 0), dt_time(12, 30))],
    'Internasional': [(dt_time(13, 0), dt_time(14, 0))]
}

ISTIRAHAT = [
    (dt_time(12, 0), dt_time(13, 0)),
    (dt_time(18, 0), dt_time(19, 0))
]

DURASI_SKS = 50
WARNA_KELAS = {
    'Online': '#4e79a7',
    'Offline': '#f28e2b',
    'Karyawan': '#59a14f',
    'Sabtu': '#e15759',
    'Internasional': '#edc948',
    'AI': '#76b7b2',
    'software': '#59a14f',
    'cybersecurity': '#e15759'
}

# ========== FUNGSI UTILITAS ==========
def parse_time(time_str):
    if isinstance(time_str, str):
        try:
            if len(time_str.split(':')) == 2:
                return datetime.strptime(time_str, '%H:%M').time()
            return datetime.strptime(time_str, '%H:%M:%S').time()
        except ValueError:
            return dt_time(8, 0)  # Default jika parsing gagal
    return time_str

def load_data():
    try:
        if not os.path.exists("data.xlsx"):
            st.error("File data.xlsx tidak ditemukan!")
            return None, None, None, None, None, None, None
        
        xls = pd.ExcelFile("data.xlsx")
        df_availability = pd.DataFrame()
        if "availability" in xls.sheet_names:
            df_availability = xls.parse("availability")
            df_availability['jam_mulai'] = df_availability['jam_mulai'].apply(parse_time)
            df_availability['jam_selesai'] = df_availability['jam_selesai'].apply(parse_time)
        
        # Load semua data
        df_kelas = xls.parse("Kelas")
        df_matkul = xls.parse("matakuliah")
        df_dosen = xls.parse("Dosen")
        df_dosen_matkul = xls.parse("dosen_matakuliah")
        df_hari = xls.parse("Hari")
        df_ruangan = xls.parse("ruangan")
        
        return df_kelas, df_matkul, df_dosen, df_dosen_matkul, df_hari['hari'].tolist(), df_ruangan['nama'].tolist(), df_availability
    except Exception as e:
        st.error(f"Gagal memuat data: {str(e)}")
        return None, None, None, None, None, None, None

def save_to_excel(df, sheet_name):
    try:
        if not os.path.exists("data.xlsx"):
            with pd.ExcelWriter("data.xlsx", engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        
        with pd.ExcelFile("data.xlsx") as xls:
            all_sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
        
        all_sheets[sheet_name] = df
        
        with pd.ExcelWriter("data.xlsx", engine='openpyxl') as writer:
            for sheet, data in all_sheets.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
        return True
    except Exception as e:
        st.error(f"Gagal menyimpan: {str(e)}")
        return False

def prioritize_room(ruangan_list, kelas):
    if kelas.startswith('TI'):
        prioritized = [r for r in ruangan_list if r.startswith('B4A')]
        others = [r for r in ruangan_list if not r.startswith('B4A')]
        return prioritized + others
    return ruangan_list

def generate_time_slots(jam_awal, jam_akhir, durasi_menit, hari, jenis_kelas):
    slots = []
    current_time = datetime.combine(datetime.today(), jam_awal)
    end_time = datetime.combine(datetime.today(), jam_akhir)
    durasi = timedelta(minutes=durasi_menit)
    
    # Penyesuaian khusus untuk jenis kelas
    if jenis_kelas.lower() in ['karyawan', 'reguler malam']:
        current_time = datetime.combine(datetime.today(), dt_time(19, 0))
        end_time = datetime.combine(datetime.today(), dt_time(21, 0))
    elif 'sabtu' in jenis_kelas.lower():
        global ISTIRAHAT
        ISTIRAHAT = [(dt_time(12, 0), dt_time(13, 0))]  # Hanya 1 istirahat untuk Sabtu
    
    while current_time + durasi <= end_time:
        start = current_time.time()
        end = (current_time + durasi).time()
        
        # Cek bentrok dengan waktu istirahat
        is_istirahat = any(
            start <= istirahat_start < end or istirahat_start <= start < istirahat_end 
            for istirahat_start, istirahat_end in ISTIRAHAT
        )
        
        if is_istirahat:
            current_time += timedelta(minutes=60)
            continue
        
        # Cek bentrok dengan waktu khusus
        if hari in WAKTU_TIDAK_BOLEH:
            is_waktu_khusus = any(
                start <= waktu_start < end or waktu_start <= start < waktu_end
                for waktu_start, waktu_end in WAKTU_TIDAK_BOLEH[hari]
            )
            if is_waktu_khusus:
                current_time += timedelta(minutes=90)
                continue
        
        slots.append((start, end))
        current_time += timedelta(minutes=durasi_menit + 10)
    
    return slots

def is_dosen_busy(nama_dosen, hari, jam_mulai, jam_selesai, df_availability, resource_tracker):
    if not df_availability.empty:
        mask = (
            (df_availability['dosen'] == nama_dosen) &
            (df_availability['hari'] == hari)
        )
        for _, row in df_availability[mask].iterrows():
            busy_start = row['jam_mulai']
            busy_end = row['jam_selesai']
            
            if not (jam_selesai <= busy_start or jam_mulai >= busy_end):
                return True
    
    for busy_start, busy_end in resource_tracker['dosen'].get((nama_dosen, hari), []):
        if not (jam_selesai <= busy_start or jam_mulai >= busy_end):
            return True
    
    return False

def is_schedule_conflict(resource_tracker, kelas, dosen, ruangan, hari, jam_mulai, jam_selesai):
    kelas_conflict = any(
        not (jam_selesai <= busy_start or jam_mulai >= busy_end)
        for busy_start, busy_end in resource_tracker['kelas'].get((kelas, hari), [])
    )
    
    dosen_conflict = any(
        not (jam_selesai <= busy_start or jam_mulai >= busy_end)
        for busy_start, busy_end in resource_tracker['dosen'].get((dosen, hari), [])
    )
    
    ruangan_conflict = False
    if ruangan and ruangan != "Zoom":
        ruangan_conflict = any(
            not (jam_selesai <= busy_start or jam_mulai >= busy_end)
            for busy_start, busy_end in resource_tracker['ruangan'].get((ruangan, hari), [])
        )
    
    return kelas_conflict or dosen_conflict or ruangan_conflict

def add_schedule(resource_tracker, kelas, dosen, ruangan, hari, jam_mulai, jam_selesai):
    resource_tracker['kelas'].setdefault((kelas, hari), []).append((jam_mulai, jam_selesai))
    resource_tracker['dosen'].setdefault((dosen, hari), []).append((jam_mulai, jam_selesai))
    if ruangan and ruangan != "Zoom":
        resource_tracker['ruangan'].setdefault((ruangan, hari), []).append((jam_mulai, jam_selesai))

# ========== FUNGSI KONSENTRASI ==========
def init_konsentrasi(df_kelas):
    """Inisialisasi kolom konsentrasi"""
    if 'konsentrasi' not in df_kelas.columns:
        df_kelas['konsentrasi'] = df_kelas['nama'].apply(
            lambda x: random.choice(KONSENTRASI_OPTIONS) if x[:4] in ['TI22', 'TI23'] else 'umum'
        )
    return df_kelas

def filter_matkul_by_konsentrasi(df_matkul, semester, konsentrasi):
    """Filter matkul berdasarkan semester dan konsentrasi"""
    matkul_semester = df_matkul[df_matkul['semester'] == semester]
    
    # Matkul umum untuk semua konsentrasi
    matkul_umum = matkul_semester[matkul_semester['Konsentrasi'] == 'umum']
    
    # Matkul khusus untuk konsentrasi tertentu
    if konsentrasi != 'umum':
        matkul_khusus = matkul_semester[
            matkul_semester['Konsentrasi'].str.contains(konsentrasi, na=False)
        ]
        return pd.concat([matkul_umum, matkul_khusus])
    return matkul_umum

def adjust_sks(matkul_df, max_sks=MAX_SKS_SEMESTER):
    """Menyesuaikan matkul agar total SKS tidak melebihi batas"""
    total_sks = matkul_df['sks'].sum()
    
    while total_sks > max_sks:
        # Prioritaskan menghapus matkul non-umum terlebih dahulu
        non_umum = matkul_df[matkul_df['Konsentrasi'] != 'umum']
        if not non_umum.empty:
            matkul_df = matkul_df.drop(non_umum.sample().index)
        else:
            # Jika semua matkul umum, hapus yang SKS-nya besar
            matkul_df = matkul_df.drop(matkul_df['sks'].idxmax())
        
        total_sks = matkul_df['sks'].sum()
    
    return matkul_df

# ========== FUNGSI PRIORITAS MATKUL ==========
def prioritize_matkul_sabtu(df_matkul):
    wajib = df_matkul[df_matkul['nama'].str.contains('lab|praktikum|jaringan', case=False, regex=True)]
    pilihan = df_matkul[~df_matkul.index.isin(wajib.index)]
    
    if len(wajib) >= 5:
        return wajib.sample(5)
    else:
        return pd.concat([wajib, pilihan.sample(min(5-len(wajib), len(pilihan)))]) 

def prioritize_matkul_karyawan(df_matkul):
    wajib = df_matkul[df_matkul['nama'].str.contains('lab|praktikum', case=False, regex=True)]
    pilihan = df_matkul[~df_matkul.index.isin(wajib.index)]
    
    if len(wajib) >= 4:
        return wajib.sample(4)
    else:
        return pd.concat([wajib, pilihan.sample(min(6-len(wajib), len(pilihan)))]) 

def prioritize_matkul_reguler_khusus(df_matkul):
    return df_matkul[df_matkul['sks'] <= 3].sample(frac=1)

def prioritize_matkul_internasional(df_matkul):
    wajib = df_matkul[df_matkul['nama'].str.contains('lab|praktikum|internasional', case=False, regex=True)]
    if len(wajib) >= 6:
        return wajib.sample(6)
    return pd.concat([wajib, df_matkul[~df_matkul.index.isin(wajib.index)].sample(min(6-len(wajib), len(df_matkul)-len(wajib)))])

# ========== FUNGSI GENERATE JADWAL UTAMA ==========
def generate_jadwal():
    df_kelas, df_matkul, df_dosen, df_dosen_matkul, df_hari, df_ruangan, df_availability = load_data()
    if df_kelas is None:
        return None

    # Inisialisasi konsentrasi
    df_kelas = init_konsentrasi(df_kelas)
    ruangan_prioritas = prioritize_room(df_ruangan, "TI")
    
    jadwal_all = []
    resource_tracker = {
        'kelas': defaultdict(list),
        'dosen': defaultdict(list),
        'ruangan': defaultdict(list)
    }

    kelas_list = df_kelas.sample(frac=1).iterrows()
    
    for _, kelas in kelas_list:
        nama_kelas = kelas['nama']
        jenis_kelas = kelas['jenis']
        konsentrasi = kelas['konsentrasi']
        prefix_kelas = nama_kelas[:4]
        jenis_kelas_suffix = nama_kelas[4:] if len(nama_kelas) > 4 else ''
        
        if prefix_kelas not in SEMESTER_KELAS:
            continue

        semester = SEMESTER_KELAS[prefix_kelas]
        
        # Filter matkul berdasarkan semester dan konsentrasi
        matkul_kelas = filter_matkul_by_konsentrasi(df_matkul, semester, konsentrasi)
        
        # Sesuaikan SKS
        matkul_kelas = adjust_sks(matkul_kelas)
        
        if matkul_kelas.empty:
            st.warning(f"Tidak ada mata kuliah untuk semester {semester}")
            continue

        # ========== PENYESUAIAN KHUSUS TIAP JENIS KELAS ==========
        if jenis_kelas_suffix == 'B':  # Kelas Sabtu
            matkul_kelas = prioritize_matkul_sabtu(matkul_kelas)
            # Untuk Sabtu, kurangi waktu istirahat
            istirahat_sabtu = [(dt_time(12, 0), dt_time(13, 0))]
        elif jenis_kelas_suffix in ['C', 'M']:  # Kelas Karyawan
            matkul_kelas = prioritize_matkul_karyawan(matkul_kelas)
        elif jenis_kelas_suffix == 'D':  # Reguler khusus
            matkul_kelas = prioritize_matkul_reguler_khusus(matkul_kelas)
        elif jenis_kelas_suffix == 'T':  # Internasional
            matkul_kelas = prioritize_matkul_internasional(matkul_kelas)
            ruangan_prioritas = [r for r in ruangan_prioritas if 'B4A' in r] + ['Zoom']

        jadwal_kelas = []

        for _, matkul in matkul_kelas.iterrows():
            # ========== FLEKSIBILITAS MODE ONLINE ==========
            must_offline = any(x in matkul['nama'].lower() for x in ['basis data', 'praktikum', 'lab', 'jaringan'])
            
            if must_offline:
                is_online = False
                ruangan_options = ruangan_prioritas.copy()
                random.shuffle(ruangan_options)
            else:
                # Atur probabilitas online berdasarkan jenis kelas
                if jenis_kelas_suffix == 'B':  # Sabtu
                    is_online = random.random() < 0.3  # 30% online
                elif jenis_kelas_suffix in ['C', 'M']:  # Karyawan
                    is_online = random.random() < 0.4  # 40% online
                elif jenis_kelas_suffix == 'D':  # Reguler khusus
                    is_online = random.random() < 0.2  # 20% online
                else:
                    is_online = matkul['Status'].lower().strip() == 'online'
                
                ruangan_options = ["Zoom"] if is_online else ruangan_prioritas.copy()
                if not is_online:
                    random.shuffle(ruangan_options)

            # ========== ATUR HARI TERSEDIA ==========
            if is_online:
                if jenis_kelas_suffix == 'B':  # Jika Sabtu tapi online, bisa weekdays
                    hari_tersedia = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Minggu']
                elif jenis_kelas_suffix in ['C', 'M']:  # Karyawan online bisa weekdays evening
                    hari_tersedia = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
                else:
                    hari_tersedia = HARI_PRIORITAS.get(jenis_kelas.lower(), ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'])
            else:
                hari_tersedia = HARI_PRIORITAS.get(jenis_kelas.lower(), ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'])

            # ========== PROSES PENJADWALAN ==========
            dosen_ids = df_dosen_matkul[df_dosen_matkul['id_matakuliah'] == matkul['id']]['id_dosen']
            dosen_tersedia = df_dosen[df_dosen['id'].isin(dosen_ids)]
            
            if dosen_tersedia.empty:
                jadwal_kelas.append({
                    'Kelas': nama_kelas,
                    'Konsentrasi': konsentrasi,
                    'Hari': 'Cek EdLink',
                    'Jam': 'Cek EdLink',
                    'Mata Kuliah': matkul['nama'],
                    'Dosen': 'Belum Ditentukan',
                    'Ruangan': 'Zoom' if is_online else 'Cek EdLink',
                    'SKS': matkul['sks'],
                    'Semester': semester,
                    'Status': 'Online' if is_online else 'Offline',
                    'Keterangan': '⚠️ Tanpa Dosen',
                    'Warna': WARNA_KELAS['Online'] if is_online else WARNA_KELAS.get(konsentrasi, WARNA_KELAS['Offline'])
                })
                continue

            scheduled = False
            attempt_log = []
            
            for attempt in range(5):
                random.shuffle(hari_tersedia)
                dosen_tersedia = dosen_tersedia.sample(frac=1)
                
                for hari in hari_tersedia:
                    if jenis_kelas_suffix == 'B' and not is_online:
                        jam_awal, jam_akhir = dt_time(8, 0), dt_time(21, 0)
                        current_istirahat = istirahat_sabtu
                    elif jenis_kelas_suffix in ['C', 'M'] and not is_online:
                        jam_awal, jam_akhir = dt_time(19, 0), dt_time(21, 0)
                        current_istirahat = []
                    else:
                        jam_awal, jam_akhir = JAM_OPERASIONAL.get(jenis_kelas.lower(), (dt_time(8, 0), dt_time(21, 0)))
                        current_istirahat = ISTIRAHAT
                    
                    possible_slots = generate_time_slots(jam_awal, jam_akhir, matkul['sks']*DURASI_SKS, hari, jenis_kelas)
                    
                    for jam_mulai, jam_selesai in possible_slots:
                        for _, dosen in dosen_tersedia.iterrows():
                            nama_dosen = dosen['nama']
                            
                            if is_dosen_busy(nama_dosen, hari, jam_mulai, jam_selesai, df_availability, resource_tracker):
                                attempt_log.append(f"Attempt {attempt}: Dosen {nama_dosen} sibuk di {hari} {jam_mulai}-{jam_selesai}")
                                continue
                                
                            if is_online:
                                if not is_schedule_conflict(resource_tracker, nama_kelas, nama_dosen, None, hari, jam_mulai, jam_selesai):
                                    add_schedule(resource_tracker, nama_kelas, nama_dosen, "Zoom", hari, jam_mulai, jam_selesai)
                                    
                                    jadwal_kelas.append({
                                        'Kelas': nama_kelas,
                                        'Konsentrasi': konsentrasi,
                                        'Hari': hari,
                                        'Jam': f"{jam_mulai.strftime('%H:%M')}-{jam_selesai.strftime('%H:%M')}",
                                        'Mata Kuliah': matkul['nama'],
                                        'Dosen': nama_dosen,
                                        'Ruangan': "Zoom",
                                        'SKS': matkul['sks'],
                                        'Semester': semester,
                                        'Status': 'Online',
                                        'Keterangan': '✅',
                                        'Warna': WARNA_KELAS['Online']
                                    })
                                    scheduled = True
                                    break
                            else:
                                for ruangan in ruangan_options:
                                    if not is_schedule_conflict(resource_tracker, nama_kelas, nama_dosen, ruangan, hari, jam_mulai, jam_selesai):
                                        add_schedule(resource_tracker, nama_kelas, nama_dosen, ruangan, hari, jam_mulai, jam_selesai)
                                        
                                        jadwal_kelas.append({
                                            'Kelas': nama_kelas,
                                            'Konsentrasi': konsentrasi,
                                            'Hari': hari,
                                            'Jam': f"{jam_mulai.strftime('%H:%M')}-{jam_selesai.strftime('%H:%M')}",
                                            'Mata Kuliah': matkul['nama'],
                                            'Dosen': nama_dosen,
                                            'Ruangan': ruangan,
                                            'SKS': matkul['sks'],
                                            'Semester': semester,
                                            'Status': 'Offline',
                                            'Keterangan': '✅',
                                            'Warna': WARNA_KELAS.get(jenis_kelas_suffix, WARNA_KELAS.get(konsentrasi, WARNA_KELAS['Offline']))
                                        })
                                        scheduled = True
                                        break
                        
                            if scheduled:
                                break
                        if scheduled:
                            break
                    if scheduled:
                        break
                if scheduled:
                    break
            
            if not scheduled and not is_online:
                for hari in hari_tersedia:
                    jam_awal, jam_akhir = JAM_OPERASIONAL.get(jenis_kelas.lower(), (dt_time(8, 0), dt_time(21, 0)))
                    possible_slots = generate_time_slots(jam_awal, jam_akhir, matkul['sks']*DURASI_SKS, hari, jenis_kelas)
                    
                    for jam_mulai, jam_selesai in possible_slots:
                        for _, dosen in dosen_tersedia.iterrows():
                            nama_dosen = dosen['nama']
                            
                            if not is_dosen_busy(nama_dosen, hari, jam_mulai, jam_selesai, df_availability, resource_tracker):
                                if not is_schedule_conflict(resource_tracker, nama_kelas, nama_dosen, None, hari, jam_mulai, jam_selesai):
                                    add_schedule(resource_tracker, nama_kelas, nama_dosen, "Zoom", hari, jam_mulai, jam_selesai)
                                    
                                    jadwal_kelas.append({
                                        'Kelas': nama_kelas,
                                        'Konsentrasi': konsentrasi,
                                        'Hari': hari,
                                        'Jam': f"{jam_mulai.strftime('%H:%M')}-{jam_selesai.strftime('%H:%M')}",
                                        'Mata Kuliah': matkul['nama'],
                                        'Dosen': nama_dosen,
                                        'Ruangan': "Zoom",
                                        'SKS': matkul['sks'],
                                        'Semester': semester,
                                        'Status': 'Online',
                                        'Keterangan': '⚠️ Auto-online: Konflik ruangan',
                                        'Warna': WARNA_KELAS['Online']
                                    })
                                    scheduled = True
                                    break
                        if scheduled:
                            break
                    if scheduled:
                        break
            
            if not scheduled:
                jadwal_kelas.append({
                    'Kelas': nama_kelas,
                    'Konsentrasi': konsentrasi,
                    'Hari': 'Cek EdLink',
                    'Jam': 'Cek EdLink',
                    'Mata Kuliah': matkul['nama'],
                    'Dosen': random.choice(dosen_tersedia['nama'].tolist()),
                    'Ruangan': 'Zoom' if is_online else 'Cek EdLink',
                    'SKS': matkul['sks'],
                    'Semester': semester,
                    'Status': 'Online' if is_online else 'Offline',
                    'Keterangan': f'⚠️ Gagal setelah 5x attempt\n' + "\n".join(attempt_log[-3:]),
                    'Warna': WARNA_KELAS['Online'] if is_online else WARNA_KELAS.get(konsentrasi, WARNA_KELAS['Offline'])
                })

        jadwal_all.extend(jadwal_kelas)
    
    return pd.DataFrame(jadwal_all)

# ========== FUNGSI KALENDER ==========
def jadwal_to_calendar_events(jadwal_df):
    events = []
    
    for _, row in jadwal_df.iterrows():
        if row['Hari'] == 'Cek EdLink':
            continue
            
        hari_to_num = {'Senin': 0, 'Selasa': 1, 'Rabu': 2, 'Kamis': 3, 
                      'Jumat': 4, 'Sabtu': 5, 'Minggu': 6}
        
        try:
            start_time = datetime.strptime(row['Jam'].split('-')[0], '%H:%M').time()
            end_time = datetime.strptime(row['Jam'].split('-')[1], '%H:%M').time()
            
            day_number = 2 + hari_to_num[row['Hari']]
            events.append({
                'title': f"{row['Mata Kuliah']} ({row['Kelas']})",
                'start': f"2023-01-{day_number:02d}T{start_time.strftime('%H:%M:%S')}",
                'end': f"2023-01-{day_number:02d}T{end_time.strftime('%H:%M:%S')}",
                'color': row['Warna'],
                'extendedProps': {
                    'dosen': row['Dosen'],
                    'sks': row['SKS'],
                    'status': row['Status'],
                    'ruangan': row['Ruangan'],
                    'keterangan': row['Keterangan'],
                    'konsentrasi': row.get('Konsentrasi', 'umum')
                }
            })
        except Exception as e:
            st.warning(f"Gagal memproses jadwal: {row['Mata Kuliah']}. Error: {str(e)}")
    
    return events

def show_calendar_view(jadwal_df):
    if calendar is None:
        st.error("Fitur kalender membutuhkan package streamlit-calendar. Install dengan: pip install streamlit-calendar")
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

# ========== TAMPILAN STREAMLIT ==========
def main():
    st.set_page_config(layout="wide", page_title="Sistem Penjadwalan Kuliah TI", page_icon="🎓")

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
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.image("https://via.placeholder.com/150x50.png?text=UNIVERSITAS", width=150)
        st.title("Menu")
        
        menu_option = st.radio(
            "Pilihan Menu",
            ["🏠 Beranda", "📅 Generate Jadwal", "👨‍🏫 Manajemen Dosen", "⚙️ Kelola Konsentrasi", "🗓️ Kalender Interaktif"],
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
        """)
        
        st.info("""
        **Fitur Unggulan:**
        1. Penjadwalan cerdas untuk berbagai jenis kelas (Reguler, Sabtu, Karyawan, Internasional)
        2. Pembagian matkul berdasarkan konsentrasi (AI, Software, Cybersecurity)
        3. Pembatasan SKS maksimal per semester (maks 21 SKS)
        4. Fleksibilitas mode online/offline
        5. Prioritas ruangan dan dosen
        6. Kalender interaktif dengan detail lengkap
        """)

    elif menu_option == "📅 Generate Jadwal":
        st.title("📅 Generate Jadwal Kuliah")
        
        df_kelas, _, _, df_dosen_matkul, _, _, _ = load_data()
        errors = []
        if df_kelas is None:
            errors.append("❌ File data.xlsx tidak ditemukan!")
        if df_dosen_matkul.empty:
            errors.append("❌ Tidak ada hubungan dosen-matkul!")
        
        if errors:
            st.error("\n".join(errors))
        else:
            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button("🔄 Generate Jadwal Baru", type="primary", use_container_width=True):
                    with st.spinner("Membuat jadwal..."):
                        st.session_state.jadwal_df = generate_jadwal()
                        st.toast("Jadwal berhasil dibuat!", icon="✅")
            
            with col2:
                if st.button("🔄 Reset Jadwal", type="secondary", use_container_width=True):
                    st.session_state.jadwal_df = None
                    st.rerun()
        
        if st.session_state.jadwal_df is not None:
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
        
        df_dosen = load_data()[2]
        df_matkul = load_data()[1]
        df_dosen_matkul = load_data()[3]
        
        with tab1:
            st.subheader("Tambah Dosen Baru")
            with st.form("form_dosen"):
                nama = st.text_input("Nama Dosen*")
                submitted = st.form_submit_button("Simpan")
                
                if submitted:
                    if not nama:
                        st.error("Nama wajib diisi!")
                    elif nama in df_dosen['nama'].values:
                        st.warning("⚠️ Dosen sudah terdaftar!")
                    else:
                        new_id = max(df_dosen['id']) + 1 if not df_dosen.empty else 1
                        new_dosen = pd.DataFrame([{'id': new_id, 'nama': nama}])
                        updated_dosen = pd.concat([df_dosen, new_dosen], ignore_index=True)
                        if save_to_excel(updated_dosen, "Dosen"):
                            st.success("Dosen berhasil ditambahkan!")
                            st.rerun()
            
            st.divider()
            st.subheader("Daftar Dosen")
            st.dataframe(df_dosen, use_container_width=True, hide_index=True)
            
            if not df_dosen.empty:
                selected_dosen = st.selectbox("Pilih Dosen untuk Dihapus", df_dosen['nama'])
                if st.button("🗑️ Hapus Dosen", type="secondary"):
                    updated_dosen = df_dosen[df_dosen['nama'] != selected_dosen]
                    if save_to_excel(updated_dosen, "Dosen"):
                        st.success("Dosen dihapus!")
                        st.rerun()
        
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
                    elif nama in df_matkul['nama'].values:
                        st.warning("⚠️ Matakuliah sudah terdaftar!")
                    else:
                        new_id = max(df_matkul['id']) + 1 if not df_matkul.empty else 1
                        new_matkul = pd.DataFrame([{
                            'id': new_id, 'nama': nama, 'sks': sks, 
                            'semester': semester, 'Status': status,
                            'Konsentrasi': konsentrasi
                        }])
                        updated_matkul = pd.concat([df_matkul, new_matkul], ignore_index=True)
                        if save_to_excel(updated_matkul, "matakuliah"):
                            st.success("Matakuliah berhasil ditambahkan!")
                            st.rerun()
            
            st.divider()
            st.subheader("Daftar Matakuliah")
            st.dataframe(df_matkul, use_container_width=True, hide_index=True)
        
        with tab3:
            st.subheader("Hubungkan Dosen dengan Matakuliah")
            col1, col2 = st.columns(2)
            with col1:
                selected_dosen = st.selectbox("Pilih Dosen", df_dosen['nama'])
                dosen_id = df_dosen[df_dosen['nama'] == selected_dosen]['id'].iloc[0] if not df_dosen.empty else None
            with col2:
                selected_matkul = st.selectbox("Pilih Matakuliah", df_matkul['nama'])
                matkul_id = df_matkul[df_matkul['nama'] == selected_matkul]['id'].iloc[0] if not df_matkul.empty else None
            
            if st.button("🔗 Hubungkan"):
                if dosen_id is None or matkul_id is None:
                    st.error("Pilih dosen dan matakuliah yang valid!")
                else:
                    if not df_dosen_matkul.empty and ((df_dosen_matkul['id_dosen'] == dosen_id) & 
                                                     (df_dosen_matkul['id_matakuliah'] == matkul_id)).any():
                        st.warning("⚠️ Hubungan sudah ada!")
                    else:
                        new_link = pd.DataFrame([{'id_dosen': dosen_id, 'id_matakuliah': matkul_id}])
                        updated_link = pd.concat([df_dosen_matkul, new_link], ignore_index=True)
                        if save_to_excel(updated_link, "dosen_matakuliah"):
                            st.success("Berhasil dihubungkan!")
                            st.rerun()
            
            st.divider()
            st.subheader("Daftar Pengajaran")
            if not df_dosen_matkul.empty:
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
                lambda x: random.choice(KONSENTRASI_OPTIONS) if x[:4] in ['TI22', 'TI23'] else 'umum'
            )
        
        st.subheader("Edit Konsentrasi Kelas")
        edited_df = st.data_editor(
            df_kelas[['nama', 'jenis', 'konsentrasi']],
            column_config={
                "konsentrasi": st.column_config.SelectboxColumn(
                    "Konsentrasi",
                    options=['umum'] + KONSENTRASI_OPTIONS,
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
        
        if st.session_state.jadwal_df is None:
            st.warning("Generate jadwal terlebih dahulu")
        else:
            show_calendar_view(st.session_state.jadwal_df)

    if st.sidebar.checkbox("🔍 Tampilkan Data Mentah"):
        st.title("Data Mentah")
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Kelas", "Mata Kuliah", "Dosen", "Ruangan", "Ketersediaan"])
        with tab1:
            df_kelas = load_data()[0]
            st.dataframe(df_kelas, use_container_width=True)
        with tab2:
            df_matkul = load_data()[1]
            st.dataframe(df_matkul, use_container_width=True)
        with tab3:
            df_dosen = load_data()[2]
            st.dataframe(df_dosen, use_container_width=True)
        with tab4:
            df_ruangan = load_data()[5]
            st.dataframe(df_ruangan, use_container_width=True)
        with tab5:
            df_availability = load_data()[6]
            st.dataframe(df_availability, use_container_width=True)

if __name__ == "__main__":
    main()