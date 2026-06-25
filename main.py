import os
import sys
import time
import subprocess
from datetime import datetime
import win32com.client as win32
from PIL import ImageGrab
import glob
import json

# Path untuk file reports (folder File report)
WORK_DIR = r"C:\Users\fresn\OneDrive\Documents\Work"

# Path untuk scripts (folder ini - tempat main.py berada)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def find_siso_file():
    """Cari file SISO OPA JAVA di folder File report (abaikan file temporary dengan ~$)"""
    print("🔄 Mencari file SISO OPA JAVA di folder File report...")

    # Cari file yang mengandung "SISO OPA JAVA"
    xlsx_files = glob.glob(os.path.join(WORK_DIR, "*SISO OPA JAVA*.xlsx"))

    # Filter: hapus file temporary yang dimulai dengan ~$
    xlsx_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]

    if not xlsx_files:
        raise FileNotFoundError("File SISO OPA JAVA tidak ditemukan di folder File report (atau semua file sedang terbuka)")

    # Ambil file terbaru jika ada multiple
    latest_file = max(xlsx_files, key=os.path.getmtime)
    print(f"   ✓ File ditemukan: {os.path.basename(latest_file)}")
    return latest_file

EXCEL_FILE_PATH = find_siso_file()


def validate_date_before_send(wb):
    """
    Validasi tanggal pada sheet Summary A1 sebelum mengirim report ke WA
    
    Logika:
    - Tanggal 2-19: today harus H-1 (1 hari setelah tanggal)
    - Tanggal 20-31: today harus H-1 (1 hari setelah tanggal)
    - Tanggal 1 (spesial): H-1 dari bulan lalu (28, 29, 30, 31 - flexible)
    
    Returns: (valid, error_message)
    """
    print("\n" + "="*60)
    print("🔍 [3.5/5] VALIDASI TANGGAL SEBELUM KIRIM KE WA")
    print("="*60)
    
    try:
        ws_summary = wb.Sheets("Summary")
        print("   ✓ Sheet 'Summary' ditemukan")
    except Exception as e:
        error_msg = f"Sheet 'Summary' tidak ditemukan: {str(e)}"
        print(f"   ❌ {error_msg}")
        return False, error_msg
    
    try:
        # Baca nilai dari A1
        date_value = ws_summary.Cells(1, 1).Value
        
        if date_value is None:
            error_msg = "Cell A1 kosong (tidak ada tanggal)"
            print(f"   ❌ {error_msg}")
            return False, error_msg
        
        # Parse tanggal dari A1
        try:
            date_num = int(float(date_value))
        except:
            error_msg = f"Cell A1 bukan angka: {date_value}"
            print(f"   ❌ {error_msg}")
            return False, error_msg
        
        if date_num < 1 or date_num > 31:
            error_msg = f"Tanggal di A1 tidak valid (harus 1-31): {date_num}"
            print(f"   ❌ {error_msg}")
            return False, error_msg
        
        print(f"   ✓ Tanggal di A1: {date_num}")
        
        # Dapatkan tanggal hari ini
        today = datetime.now()
        today_day = today.day
        today_month = today.month
        today_year = today.year
        
        print(f"   📅 Hari ini: {today.strftime('%Y-%m-%d (%A)')} (Tanggal: {today_day})")
        
        # Tentukan expected_day berdasarkan tanggal di A1
        expected_day = None
        expected_month = today_month
        expected_year = today_year
        schedule_type = None
        validation_result = False
        
        if date_num == 1:
            # Tanggal 1: H-1 adalah bulan lalu (28, 29, 30, atau 31)
            # Check apakah kemarin adalah hari terakhir bulan lalu
            from datetime import timedelta
            yesterday_full = today - timedelta(days=1)
            
            if yesterday_full.month != today_month:
                # Kemarin adalah bulan lalu - ini benar!
                # Accept tanggal 28-31 dari bulan lalu
                yesterday_day = yesterday_full.day
                if 28 <= yesterday_day <= 31:
                    validation_result = True
                    schedule_type = f"Tanggal 1 (H-1 dari bulan lalu: tanggal {yesterday_day})"
                    print(f"   ✓ Status: Tanggal 1 - kemarin (H-1) adalah {yesterday_day} bulan lalu ✅")
                else:
                    validation_result = False
                    schedule_type = "Tanggal 1"
                    print(f"   ⚠️ Status: Kemarin bulan lalu tapi bukan 28-31 (tanggal {yesterday_day})")
            else:
                validation_result = False
                schedule_type = "Tanggal 1"
                print(f"   ⚠️ Status: Belum awal bulan (kemarin masih bulan yang sama)")
        
        elif 2 <= date_num <= 31:
            # Tanggal 2-31: H-1 (1 hari setelah tanggal) - semua H-1
            expected_day = date_num + 1
            schedule_type = f"Tanggal {date_num} (H-1)"
            
            from calendar import monthrange
            days_in_month = monthrange(today_year, today_month)[1]
            
            if expected_day > days_in_month:
                # Overflow ke bulan berikutnya (biasanya ke tanggal 1)
                expected_day = expected_day - days_in_month
                expected_month = today_month + 1
                if expected_month > 12:
                    expected_month = 1
                    expected_year = today_year + 1
            
            validation_result = (today_day == expected_day and today_month == expected_month)
            print(f"   📋 Expected: Tanggal {expected_day} bulan {expected_month} (H-1 dari {date_num})")
        
        # Tampilkan hasil validasi
        if validation_result:
            print(f"\n   ✅ VALIDASI BERHASIL!")
            print(f"   📅 Jadwal: {schedule_type}")
            print(f"   📊 Report: Mengirim data H-1 dari tanggal {date_num}")
            return True, None
        else:
            error_msg = f"❌ TANGGAL TIDAK SESUAI!\n"
            error_msg += f"   - A1 menunjukkan: {date_num}\n"
            error_msg += f"   - Jadwal: {schedule_type}\n"
            if date_num == 1:
                error_msg += f"   - Expected: Hari ini harus awal bulan, kemarin harus 28-31 bulan lalu\n"
            else:
                error_msg += f"   - Expected: Tanggal {expected_day} (hari ini: {today_day})\n"
            error_msg += f"   - Mungkin file belum di-update atau tanggal A1 salah"
            print(f"\n   {error_msg}")
            return False, error_msg
    
    except Exception as e:
        error_msg = f"Error saat validasi tanggal: {str(e)}"
        print(f"   ❌ {error_msg}")
        import traceback
        traceback.print_exc()
        return False, error_msg


def send_error_to_wa(error_message):
    """Kirim pesan error ke WhatsApp (connect OCBC gagal)"""
    try:
        print(f"\n⚠️ Mengirim notifikasi error ke WhatsApp...")
        print(f"   Pesan: connect OCBC gagal")
        print(f"   Detail: {error_message[:100]}")
        
        error_config = {
            'group_ids': [],
            'caption': f"❌ PROSES GAGAL\nAlasan: connect OCBC gagal\nDetail: {error_message}",
            'screenshot_count': 0,
            'is_error': True
        }
        
        config_file = os.path.join(SCRIPT_DIR, 'send_config.json')
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(error_config, f, ensure_ascii=False, indent=2)
        
        # Jalankan wa_bot.js untuk kirim error
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        
        result = subprocess.run(
            ["node", os.path.join(SCRIPT_DIR, "wa_bot.js")],
            cwd=SCRIPT_DIR,
            timeout=60,
            encoding="utf-8",
            errors="replace",
            env=env
        )
        
        # Cleanup
        if os.path.exists(config_file):
            os.remove(config_file)
        
        print(f"   ✓ Notifikasi error terkirim")
    
    except Exception as e:
        print(f"   ⚠️ Gagal mengirim notifikasi error: {str(e)}")


def parse_group_mapping(wb):
    """Parse sheet Group untuk mendapatkan mapping Group Name -> Group IDs"""
    print("\n🔍 Membaca sheet Group untuk mapping ID Grup...")
    
    try:
        ws_group = wb.Sheets("Group")
    except Exception as e:
        print(f"   ⚠️ Sheet 'Group' tidak ditemukan: {str(e)}")
        return {}
    
    # Cari header: Group, Group ID
    headers = {}
    header_row = None
    
    for search_row in range(1, 20):
        cell_value = ws_group.Cells(search_row, 1).Value
        if cell_value and str(cell_value).strip().lower() == 'group':
            header_row = search_row
            print(f"   ✓ Header ditemukan di Row {header_row}")
            
            for col in range(1, 10):
                header_text = ws_group.Cells(header_row, col).Value
                if header_text:
                    headers[header_text.strip().lower()] = col
            break
    
    if header_row is None:
        print("   ⚠️ Header 'Group' tidak ditemukan di sheet Group")
        return {}
    
    print(f"   📋 Header: {headers}")
    
    # Baca data group mapping
    group_mapping = {}
    row = header_row + 1
    
    print(f"\n🔄 Membaca mapping Group:")
    while True:
        group_name = ws_group.Cells(row, 1).Value
        if group_name is None or str(group_name).strip() == '':
            break
        
        group_id = None
        if 'group id' in headers:
            group_id = ws_group.Cells(row, headers['group id']).Value
        
        if group_id:
            group_name_clean = str(group_name).strip()
            group_id_clean = str(group_id).strip()
            group_mapping[group_name_clean] = group_id_clean
            print(f"   ✓ {group_name_clean} → {group_id_clean}")
        
        row += 1
    
    print(f"\n   ✓ {len(group_mapping)} grup mapping ditemukan")
    return group_mapping


def parse_caption_table(ws_caption):
    """Parse sheet Caption:
       1. Cari header (Dashboard Name, Sheet, From, To, Group) - bisa di row mana saja
       2. Baca list dashboard hingga baris kosong
       3. Setelah baris kosong, cari "Caption" header
       4. Baca isi caption di bawahnya sampai baris kosong
    """

    # Langkah 1: Cari header untuk dashboard list (cari di semua row)
    print("\n🔍 Mencari header dashboard...")
    headers = {}
    header_row = None

    for search_row in range(1, 20):  # Cari di row 1-20
        cell_value = ws_caption.Cells(search_row, 1).Value  # Kolom A
        if cell_value and str(cell_value).strip().lower() == 'dashboard name':
            header_row = search_row
            print(f"   ✓ Header ditemukan di Row {header_row}")

            # Baca semua header di row ini
            for col in range(1, 10):
                header_text = ws_caption.Cells(header_row, col).Value
                if header_text:
                    headers[header_text.strip().lower()] = col
            break

    if header_row is None:
        # Cari alternatif: "dashboard"
        for search_row in range(1, 20):
            cell_value = ws_caption.Cells(search_row, 1).Value
            if cell_value and 'dashboard' in str(cell_value).strip().lower():
                header_row = search_row
                print(f"   ✓ Header ditemukan di Row {header_row} (partial match)")

                for col in range(1, 10):
                    header_text = ws_caption.Cells(header_row, col).Value
                    if header_text:
                        headers[header_text.strip().lower()] = col
                break

    if header_row is None:
        raise ValueError("Tidak bisa menemukan header 'Dashboard Name' di sheet Caption")

    print(f"   📋 Header: {headers}")

    # Langkah 2: Baca data dashboard
    data_list = []
    row = header_row + 1
    last_data_row = header_row

    print(f"\n🔄 Membaca data dashboard mulai dari Row {row}:")
    while True:
        # Cek kolom yang ada - gunakan kolom pertama yang ada header
        first_header_col = min(headers.values()) if headers else 1
        dashboard_name = ws_caption.Cells(row, first_header_col).Value

        if dashboard_name is None or str(dashboard_name).strip() == '':
            # Baris kosong - pemisah
            print(f"   ⚠️ Baris kosong di Row {row} (pemisah)")
            last_data_row = row
            break

        # Baca data dengan flexible column checking
        sheet_name = None
        from_range = None
        to_range = None
        group_name = None

        if 'sheet' in headers:
            sheet_name = ws_caption.Cells(row, headers['sheet']).Value
        if 'from' in headers:
            from_range = ws_caption.Cells(row, headers['from']).Value
        if 'to' in headers:
            to_range = ws_caption.Cells(row, headers['to']).Value
        if 'group' in headers:
            group_name = ws_caption.Cells(row, headers['group']).Value

        if sheet_name and from_range and to_range:
            data_list.append({
                'dashboard_name': str(dashboard_name).strip(),
                'sheet': str(sheet_name).strip(),
                'from': str(from_range).strip(),
                'to': str(to_range).strip(),
                'group': str(group_name).strip() if group_name else 'Default',
            })
            print(f"   ✓ Row {row}: {dashboard_name} | Sheet: {sheet_name} | Range: {from_range}:{to_range} | Group: {str(group_name).strip() if group_name else 'Default'}")

        row += 1

    # Langkah 3: Cari "Caption" header setelah baris kosong
    print(f"\n🔍 Mencari section 'Caption'...")
    caption_header_row = None
    search_row = last_data_row + 1

    while search_row < last_data_row + 10:
        cell_value = ws_caption.Cells(search_row, 1).Value
        if cell_value and str(cell_value).strip().lower() == 'caption':
            caption_header_row = search_row
            print(f"   ✓ Header 'Caption' ditemukan di Row {caption_header_row}")
            break
        search_row += 1

    # Langkah 4: Baca isi caption
    caption_text = ""
    if caption_header_row is not None:
        caption_lines = []
        caption_row = caption_header_row + 1

        print(f"\n📝 Membaca isi caption mulai dari Row {caption_row}:")
        while True:
            cell_value = ws_caption.Cells(caption_row, 1).Value
            if cell_value is None or str(cell_value).strip() == '':
                print(f"   ⚠️ Baris kosong di Row {caption_row} (end of caption)")
                break

            caption_text_part = str(cell_value).strip()
            caption_lines.append(caption_text_part)
            print(f"   ✓ Row {caption_row}: {caption_text_part}")
            caption_row += 1

        caption_text = "\n".join(caption_lines)
    else:
        print("   ⚠️ Section 'Caption' tidak ditemukan")

    return data_list, caption_text


def group_dashboards_by_consecutive_group(capture_config):
    """Group consecutive dashboards dengan Group yang sama"""
    if not capture_config:
        return []
    
    grouped = []
    current_group = None
    current_dashboards = []
    current_caption_lines = []
    
    for i, config in enumerate(capture_config):
        dashboard_group = config.get('group', 'Default')
        
        # Jika group berubah, simpan group sebelumnya
        if dashboard_group != current_group and current_dashboards:
            grouped.append({
                'group': current_group,
                'dashboards': current_dashboards,
                'caption_lines': current_caption_lines
            })
            current_dashboards = []
            current_caption_lines = []
        
        current_group = dashboard_group
        current_dashboards.append(config)
    
    # Simpan group terakhir
    if current_dashboards:
        grouped.append({
            'group': current_group,
            'dashboards': current_dashboards,
            'caption_lines': current_caption_lines
        })
    
    return grouped


def take_all_screenshots(wb, excel, group_mapping, all_dashboards):
    """
    Ambil semua screenshots dengan naming based on group.
    
    Returns:
    - screenshot_map: {group_name: [list of image_paths]}
    - group_to_ids: {group_name: group_ids_list}
    """
    print("\n" + "="*60)
    print("🔄 [5/5] MENGAMBIL SEMUA SCREENSHOTS")
    print("="*60)
    
    screenshot_map = {}  # {group_name: [image_paths]}
    group_to_ids = {}    # {group_name: group_ids_list}
    group_counters = {}  # {group_name: counter} untuk naming
    
    total_screenshots = 0
    
    for group_name, dashboards in all_dashboards.items():
        print(f"\n📸 GROUP: {group_name}")
        print(f"   Dashboards: {len(dashboards)} item")
        
        # Cari Group IDs dari mapping
        group_ids = group_mapping.get(group_name, '')
        if not group_ids:
            print(f"   ⚠️ Group '{group_name}' tidak ditemukan di sheet Group - SKIP")
            continue
        
        # Group ID bisa multiple, dipisah dengan ;
        group_ids_list = [gid.strip() for gid in group_ids.split(';') if gid.strip()]
        group_to_ids[group_name] = group_ids_list
        print(f"   📍 Group IDs: {', '.join(group_ids_list)}")
        
        # Initialize counter untuk group ini
        group_counters[group_name] = 0
        screenshot_map[group_name] = []
        
        # Ambil screenshot untuk setiap dashboard
        for dash_idx, config in enumerate(dashboards, 1):
            try:
                sheet_name = config['sheet']
                from_range = config['from']
                to_range = config['to']
                dashboard_name = config['dashboard_name']

                print(f"\n   📸 [{dash_idx}/{len(dashboards)}] {dashboard_name}")
                print(f"      Sheet: {sheet_name} | Range: {from_range}:{to_range}")

                ws = wb.Sheets(sheet_name)
                ws.Activate()
                time.sleep(0.5)

                range_address = f"{from_range}:{to_range}"
                selected_range = ws.Range(range_address)
                selected_range.Select()
                time.sleep(0.3)

                selected_range.Copy()
                time.sleep(0.5)

                img = ImageGrab.grabclipboard()
                if img:
                    group_counters[group_name] += 1
                    img = img.crop(img.getbbox())

                    # ✅ NAMING: group_name_counter.jpg (e.g., "IM3_1.jpg", "3ID_2.jpg")
                    image_filename = f"{group_name}_{group_counters[group_name]}.jpg"
                    image_path = os.path.join(SCRIPT_DIR, image_filename)

                    img_rgb = img.convert('RGB')
                    img_rgb.save(image_path, 'JPEG', quality=95)
                    screenshot_map[group_name].append(image_path)
                    total_screenshots += 1
                    print(f"      ✓ Screenshot berhasil: {image_filename} ({img_rgb.size})")

                    excel.SendKeys('{ESCAPE}')
                else:
                    print(f"      ⚠️ Clipboard kosong untuk {dashboard_name}")
                    excel.SendKeys('{ESCAPE}')

            except Exception as e:
                print(f"      ❌ Error capturing {config['dashboard_name']}: {str(e)[:80]}")
                excel.SendKeys('{ESCAPE}')
                continue
        
        print(f"   ✅ {group_name}: {group_counters[group_name]} screenshot berhasil")
    
    print(f"\n✅ Total screenshots: {total_screenshots}")
    return screenshot_map, group_to_ids


def wait_for_connections_refresh(wb, max_wait=150, poll_interval=3):
    """
    ✅ IMPROVED: Polling untuk detect kapan query GCP selesai.
    
    Smart Fallback Strategy:
    1. Primary: Check CalculationState = -1 (Done) → return immediately
    2. Fallback: Jika state tetap Calculating/Pending selama 45 detik tanpa perubahan,
       dianggap query sudah selesai (Excel UI mungkin tidak update, tapi data sudah ready)
    3. Max timeout: 1800s (30 menit)
    
    Ini akan STOP waiting segera setelah query benar-benar selesai, tidak akan nunggu 1800s!
    """
    print(f"🔄 [3/5] Menunggu query GCP selesai (max {max_wait}s)...")
    print(f"   💡 Smart fallback: jika state stabil 45s, dianggap selesai")
    
    start = time.time()
    elapsed = 0
    stable_count = 0
    last_state = None
    last_state_change_time = time.time()
    consecutive_errors = 0
    no_change_threshold = 45  # 45 detik state tidak berubah = selesai
    
    while elapsed < max_wait:
        try:
            # Force calculate
            excel_app = wb.Application
            excel_app.Calculate()
            
            try:
                # Check state calculation
                calc_state = excel_app.CalculationState
                # -1 = Done, 1 = Calculating, 2 = Pending
                state_name = {-1: "✅ Done", 1: "🔄 Calculating", 2: "⏳ Pending"}.get(calc_state, "❓ Unknown")
                
                # ✅ PRIMARY: Jika state = Done, return immediately
                if calc_state == -1:
                    stable_count += 1
                    print(f"   {state_name} ({stable_count}/3 confirm) ... {elapsed}s")
                    
                    if stable_count >= 3:
                        print(f"   ✅ Query selesai (CalculationState = Done)! ({elapsed}s)")
                        time.sleep(2)
                        return True
                else:
                    # Still Calculating or Pending
                    stable_count = 0
                    
                    # ✅ FALLBACK: Check apakah state berubah atau tetap sama?
                    if calc_state != last_state:
                        # State berubah - reset timer
                        last_state = calc_state
                        last_state_change_time = time.time()
                        print(f"   {state_name} ... {elapsed}s / {max_wait}s [state changed]")
                    else:
                        # State tetap sama
                        time_since_change = int(time.time() - last_state_change_time)
                        print(f"   {state_name} ... {elapsed}s / {max_wait}s [stable: {time_since_change}s]")
                        
                        # Jika state tidak berubah selama 45 detik, assume selesai
                        # (mungkin Excel UI tidak update tapi data sudah ready)
                        if time_since_change >= no_change_threshold:
                            print(f"   ✅ State stabil {no_change_threshold}s, query dianggap selesai! ({elapsed}s)")
                            time.sleep(2)
                            return True
                
                consecutive_errors = 0
                
            except Exception as state_error:
                consecutive_errors += 1
                err = str(state_error)[:40]
                print(f"   ⏳ Checking status... ({elapsed}s) [error {consecutive_errors}x: {err}]")
                stable_count = 0
                
                if consecutive_errors >= 10:
                    print(f"   ⚠️ Terlalu banyak error saat cek status, assume selesai")
                    return False
        
        except Exception as e:
            err_msg = str(e)
            
            if "rejected by callee" in err_msg.lower() or "-2147418111" in err_msg:
                consecutive_errors += 1
                print(f"   ⏳ Excel busy (COM error), retry... ({elapsed}s) [{consecutive_errors}x]")
                stable_count = 0
                
                if consecutive_errors >= 10:
                    print(f"   ⚠️ Excel terus busy, assume timeout")
                    return False
            else:
                print(f"   ⚠️ Unexpected error: {err_msg[:60]}")
                stable_count = 0
        
        time.sleep(poll_interval)
        elapsed = int(time.time() - start)
    
    # Timeout reached
    print(f"\n⚠️ TIMEOUT setelah {max_wait}s")
    print(f"   ⚠️ Lanjut ke screenshot dengan data yang ada")
    return False


def refresh_dan_screenshot():
    print("\n" + "="*60)
    print("📊 MEMBUKA FILE DAN REFRESH DATA GCP")
    print("="*60)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.Interactive = True

    print("✓ Excel Application dibuat dan Visible=True")

    wb = None

    try:
        print(f"\n🔄 [1/5] Membuka file: {os.path.basename(EXCEL_FILE_PATH)}")
        print(f"    Path: {EXCEL_FILE_PATH}")
        print("   ⏳ Opening workbook...")
        
        time.sleep(2)

        start_time = time.time()
        try:
            print(f"   📂 Mencoba buka: {EXCEL_FILE_PATH}")
            print(f"   ⏳ Waiting for Open...")
            
            wb = excel.Workbooks.Open(
                Filename=EXCEL_FILE_PATH,
                UpdateLinks=False,
                ReadOnly=False
            )
            open_time = time.time() - start_time
            print(f"   ✓ File terbuka ({open_time:.1f}s)")
        except Exception as open_error:
            print(f"   ❌ Error membuka file: {str(open_error)}")
            print(f"   📝 File path: {EXCEL_FILE_PATH}")
            print(f"   📝 File exists? {os.path.exists(EXCEL_FILE_PATH)}")
            raise

        print("\n⏳ Menunggu Excel siap (5 detik)...")
        time.sleep(5)

        print("🔄 [2/5] Melakukan refresh data dari GCP...")
        max_retries = 6
        for attempt in range(1, max_retries + 1):
            try:
                wb.RefreshAll()
                print(f"   ✓ RefreshAll berhasil dipanggil (percobaan {attempt})")
                break
            except Exception as e:
                if attempt == max_retries:
                    print(f"   ❌ RefreshAll tetap gagal setelah {max_retries} percobaan: {str(e)[:100]}")
                    raise
                wait_s = attempt * 3
                print(f"   ⚠️ Excel masih busy (percobaan {attempt}/{max_retries}): {str(e)[:80]}")
                print(f"   ⏳ Retry dalam {wait_s}s...")
                time.sleep(wait_s)

        time.sleep(2)

        # ✅ SMART WAIT: akan stop as soon as query finishes, not wait full 1800s!
        wait_for_connections_refresh(wb, max_wait=150, poll_interval=3)
        print("   ✓ Proses refresh GCP selesai!")

        # 🔍 VALIDASI TANGGAL SEBELUM MENGIRIM KE WA
        is_date_valid, date_error_msg = validate_date_before_send(wb)
        
        if not is_date_valid:
            print(f"\n❌ VALIDASI GAGAL - MENGHENTIKAN PROSES")
            print(f"   Mengirimkan notifikasi error ke WhatsApp...")
            send_error_to_wa(date_error_msg)
            print(f"\n✅ Notifikasi error telah dikirim")
            print(f"   Jalankan kembali setelah tanggal di A1 di-update")
            return  # Stop di sini, jangan lanjut ke screenshot

        # Parse Group mapping dari sheet Group
        group_mapping = parse_group_mapping(wb)

        print("\n🔄 [4/5] Membaca konfigurasi dashboard dari semua sheets...")

        sheets_to_process = ["Caption", "Caption 2"]
        all_dashboards = {}  # {group_name: [list of dashboard configs]}
        all_captions = {}    # {sheet_label: caption_text}

        # STEP 1: Collect semua dashboards dari semua sheets
        for sheet_label in sheets_to_process:
            try:
                ws_caption = wb.Sheets(sheet_label)
                print(f"\n📄 Membaca sheet: {sheet_label}")

                capture_config, caption_text = parse_caption_table(ws_caption)

                if not capture_config:
                    print(f"   ⚠️ Tidak ada dashboard di sheet {sheet_label}")
                    continue

                print(f"   ✓ {len(capture_config)} dashboard ditemukan")
                all_captions[sheet_label] = caption_text

                # Debug: print caption info
                if caption_text:
                    print(f"   📝 Caption tersimpan: {len(caption_text)} karakter")
                else:
                    print(f"   ⚠️ Caption kosong untuk sheet {sheet_label}")

                # Group dashboards by group name
                for config in capture_config:
                    group_name = config.get('group', 'Default')
                    if group_name not in all_dashboards:
                        all_dashboards[group_name] = []
                    all_dashboards[group_name].append(config)

            except Exception as e:
                print(f"   ⚠️ Error membaca sheet {sheet_label}: {str(e)[:80]}")
                continue

        if not all_dashboards:
            print("\n⚠️ Tidak ada dashboard ditemukan di semua sheets")
            return

        print(f"\n✅ Total group ditemukan: {len(all_dashboards)}")
        for group_name, dashboards in all_dashboards.items():
            print(f"   - {group_name}: {len(dashboards)} dashboard")

        # STEP 2: Ambil semua screenshots dengan naming based on group
        screenshot_map, group_to_ids = take_all_screenshots(wb, excel, group_mapping, all_dashboards)

        if not screenshot_map or all(len(images) == 0 for images in screenshot_map.values()):
            print("\n❌ Tidak ada screenshot yang berhasil diambil")
            return

        # Validasi jumlah dashboard vs screenshot
        total_dashboards = sum(len(dashboards) for dashboards in all_dashboards.values())
        total_screenshots = sum(len(images) for images in screenshot_map.values())
        print("\n📊 Validasi total dashboard vs screenshot:")
        print(f"   - Total dashboard: {total_dashboards}")
        print(f"   - Total screenshot: {total_screenshots}")

        if total_dashboards != total_screenshots:
            print("\n❌ VALIDASI GAGAL: Jumlah screenshot tidak sesuai jumlah dashboard")
            for group_name in sorted(all_dashboards.keys()):
                dashboards_count = len(all_dashboards[group_name])
                screenshot_count = len(screenshot_map.get(group_name, []))
                print(f"   - Group {group_name}: dashboard={dashboards_count}, screenshot={screenshot_count}")
            print("\n⚠️ Proses kirim WA dibatalkan. Periksa kembali konfigurasi sheet Caption/Caption 2.")
            return

        print("\n✅ Validasi berhasil: jumlah screenshot sesuai total dashboard")

        # STEP 3: Send per group (images dulu, baru caption)
        print("\n" + "="*60)
        print("📱 MENGIRIM KE GRUP")
        print("="*60)

        for group_name, image_paths in screenshot_map.items():
            if not image_paths:
                print(f"\n⚠️ Group '{group_name}' tidak ada images - SKIP")
                continue

            group_ids_list = group_to_ids.get(group_name, [])
            if not group_ids_list:
                print(f"\n⚠️ Group '{group_name}' tidak ada Group IDs - SKIP")
                continue

            print(f"\n{'='*60}")
            print(f"📍 GROUP: {group_name}")
            print(f"   Images: {len(image_paths)}")
            print(f"   Target: {', '.join(group_ids_list)}")
            print(f"{'='*60}")

            # Ambil caption dari sheet Caption (utama)
            caption_text = all_captions.get('Caption', '')
            if not caption_text:
                # Fallback ke Caption 2 jika Caption kosong
                caption_text = all_captions.get('Caption 2', '')
            
            print(f"   📝 Caption untuk kirim: {len(caption_text) if caption_text else 0} karakter")
            if not caption_text:
                print(f"      ⚠️ Catatan: all_captions keys = {list(all_captions.keys())}")

            
            # Simpan caption ke file untuk referensi
            if caption_text:
                save_caption_to_file(caption_text)

            # Save config dan send
            save_send_config_with_images(caption_text, group_ids_list, image_paths)

            print(f"\n🔄 Mengirim ke WhatsApp...")
            send_wa_report()

            print(f"✅ Group {group_name} selesai!")

        # STEP 4: Cleanup semua file screenshot
        cleanup_all_screenshot_files()

    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except:
            pass

        try:
            excel.Quit()
        except:
            pass


def send_wa_report():
    """Jalankan wa_bot.js untuk mengirim screenshot ke WA"""
    try:
        print("\n🔄 Menjalankan WA Bot untuk mengirim screenshot...")

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        result = subprocess.run(
            ["node", os.path.join(SCRIPT_DIR, "wa_bot.js")],
            cwd=SCRIPT_DIR,
            timeout=300,
            encoding="utf-8",
            errors="replace",
            env=env
        )

        if result.returncode == 0:
            print("✅ WA Bot selesai!")
        else:
            print(f"❌ WA Bot gagal dengan exit code: {result.returncode}")
            print("   (lihat log di atas untuk detail error dari wa_bot.js)")
            raise RuntimeError("WA Bot gagal mengirim report")

    except subprocess.TimeoutExpired:
        print("❌ WA Bot timeout (terlalu lama)")
        raise
    except Exception as e:
        print(f"❌ Error menjalankan WA Bot: {str(e)}")
        raise


def save_send_config_with_images(caption_text, group_ids, image_paths):
    """Simpan konfigurasi untuk wa_bot.js dengan list image paths"""
    try:
        send_config = {
            'group_ids': group_ids,
            'caption': caption_text if caption_text and caption_text.strip() else '',
            'image_files': image_paths  # ✅ BARU: list of image paths
        }
        
        config_file = os.path.join(SCRIPT_DIR, 'send_config.json')
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(send_config, f, ensure_ascii=False, indent=2)
        
        print(f"\n📝 Konfigurasi pengiriman disimpan: send_config.json")
        print(f"   Group IDs: {', '.join(group_ids)}")
        print(f"   Images: {len(image_paths)} file")
        if caption_text and caption_text.strip():
            print(f"   Caption: {len(caption_text)} karakter")
        
    except Exception as e:
        print(f"⚠️ Error menyimpan konfigurasi pengiriman: {str(e)}")


def save_caption_to_file(caption_text):
    """Simpan caption ke file Caption.txt"""
    try:
        caption_file = os.path.join(SCRIPT_DIR, 'Caption.txt')
        with open(caption_file, 'w', encoding='utf-8') as f:
            f.write(caption_text)
        print(f"   ✓ Caption disimpan: Caption.txt ({len(caption_text)} karakter)")
    except Exception as e:
        print(f"   ⚠️ Error simpan caption: {str(e)}")


def cleanup_all_screenshot_files():
    """Hapus semua file screenshot jpg dan temp files"""
    try:
        print("\n🗑️ Membersihkan file screenshot temporary...")

        # Hapus semua jpg files (hasil screenshot)
        jpg_files = glob.glob(os.path.join(SCRIPT_DIR, "*.jpg"))
        for file in jpg_files:
            try:
                os.remove(file)
                print(f"   ✓ Dihapus: {os.path.basename(file)}")
            except Exception as e:
                print(f"   ⚠️ Gagal hapus {os.path.basename(file)}: {str(e)}")

        # Hapus backup/legacy files jika ada
        legacy_files = glob.glob(os.path.join(SCRIPT_DIR, "temp_report_*.png"))
        for file in legacy_files:
            try:
                os.remove(file)
                print(f"   ✓ Dihapus: {os.path.basename(file)}")
            except Exception as e:
                print(f"   ⚠️ Gagal hapus {os.path.basename(file)}: {str(e)}")

        # Hapus config
        send_config_file = os.path.join(SCRIPT_DIR, "send_config.json")
        if os.path.exists(send_config_file):
            os.remove(send_config_file)
            print(f"   ✓ Dihapus: send_config.json")

        # Hapus Caption.txt jika ada
        caption_file = os.path.join(SCRIPT_DIR, "Caption.txt")
        if os.path.exists(caption_file):
            os.remove(caption_file)
            print(f"   ✓ Dihapus: Caption.txt")

        print("✅ Cleanup selesai!")

    except Exception as e:
        print(f"⚠️ Error saat cleanup: {str(e)}")


def save_send_config(caption_text, group_ids, screenshot_count):
    """Simpan konfigurasi untuk wa_bot.js: caption text, group IDs, dan jumlah screenshots"""
    try:
        send_config = {
            'group_ids': group_ids,
            'caption': caption_text if caption_text and caption_text.strip() else '',
            'screenshot_count': screenshot_count
        }
        
        config_file = os.path.join(SCRIPT_DIR, 'send_config.json')
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(send_config, f, ensure_ascii=False, indent=2)
        
        print(f"\n📝 Konfigurasi pengiriman disimpan: send_config.json")
        print(f"   Group IDs: {', '.join(group_ids)}")
        print(f"   Screenshots: {screenshot_count} file")
        if caption_text and caption_text.strip():
            print(f"   Caption: {len(caption_text)} karakter")
        
    except Exception as e:
        print(f"⚠️ Error menyimpan konfigurasi pengiriman: {str(e)}")


if __name__ == "__main__":
    try:
        refresh_dan_screenshot()
    except FileNotFoundError as e:
        print(f"\n❌ FILE ERROR: {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ ERROR: {type(e).__name__}: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)