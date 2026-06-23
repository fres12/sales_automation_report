import os
import sys
import time
import subprocess
from datetime import datetime
import win32com.client as win32
from PIL import ImageGrab, ImageEnhance
import glob

# Path untuk file reports (folder File report)
WORK_DIR = r"C:\Users\fresn\OneDrive - Indosatooredoo Hutchison\C_Java_Analytic - SISO OPA Q3 2026"

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
    """Group consecutive dashboards dengan Group yang sama
    Input: [{'dashboard_name': '...', 'group': 'IM3', ...}, ...]
    Output: [{'group': 'IM3', 'dashboards': [...], 'caption': '...'}, ...]
    """
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


def wait_for_connections_refresh(wb, max_wait=1800, poll_interval=3):
 
    print(f"🔄 [3/5] Menunggu query GCP selesai (max {max_wait}s)...")
    
    start = time.time()
    elapsed = 0
    stable_count = 0
    last_state = None
    consecutive_errors = 0
    
    while elapsed < max_wait:
        try:
            # Force calculate - ini yang trigger query untuk jalan
            excel_app = wb.Application
            excel_app.Calculate()
            
            try:
                # Check state calculation
                calc_state = excel_app.CalculationState
                # -1 = Done, 1 = Calculating, 2 = Pending
                state_name = {-1: "✅ Done", 1: "🔄 Calculating", 2: "⏳ Pending"}.get(calc_state, "❓ Unknown")
                
                if calc_state == -1:  # DONE = query selesai
                    stable_count += 1
                    print(f"   {state_name} ({stable_count}/3 confirm) ... {elapsed}s")
                    
                    # Tunggu 3x sebelum betul-betul dianggap selesai
                    # Ini untuk avoid false positive (Excel bilang Done tapi data belum fully synced)
                    if stable_count >= 3:
                        print(f"   ✅ Query benar-benar selesai! ({elapsed}s)")
                        time.sleep(2)  # Extra buffer untuk Excel selesai write
                        return True
                else:
                    # Still calculating/pending
                    stable_count = 0
                    print(f"   {state_name} ... {elapsed}s / {max_wait}s")
                
                last_state = calc_state
                consecutive_errors = 0
                
            except Exception as state_error:
                consecutive_errors += 1
                err = str(state_error)[:40]
                print(f"   ⏳ Checking status... ({elapsed}s) [error {consecutive_errors}x: {err}]")
                stable_count = 0
                
                # Kalau error berkali-kali, anggap mungkin Excel crash
                if consecutive_errors >= 10:
                    print(f"   ⚠️ Terlalu banyak error, assume selesai atau timeout")
                    return False
        
        except Exception as e:
            # Error saat access Excel object itself
            err_msg = str(e)
            
            if "rejected by callee" in err_msg.lower() or "-2147418111" in err_msg:
                # COM error - Excel busy
                consecutive_errors += 1
                print(f"   ⏳ Excel busy (COM error), retry... ({elapsed}s) [{consecutive_errors}x]")
                stable_count = 0
                
                if consecutive_errors >= 10:
                    print(f"   ⚠️ Excel terus busy, assume timeout")
                    return False
            else:
                # Unexpected error
                print(f"   ⚠️ Unexpected error: {err_msg[:60]}")
                stable_count = 0
        
        # Wait interval sebelum polling lagi
        time.sleep(poll_interval)
        elapsed = int(time.time() - start)
    
    # Timeout reached
    print(f"\n⚠️ TIMEOUT setelah {max_wait}s")
    print(f"   ⚠️ Data mungkin belum 100% terupdate, tapi lanjut ke screenshot")
    return False


def refresh_dan_screenshot():
    print("\n" + "="*60)
    print("📊 MEMBUKA FILE DAN REFRESH DATA GCP")
    print("="*60)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.Interactive = True  # Ensure Excel is interactive
    
    print("✓ Excel Application dibuat dan Visible=True")

    wb = None

    try:
        print(f"\n🔄 [1/5] Membuka file: {os.path.basename(EXCEL_FILE_PATH)}")
        print(f"    Path: {EXCEL_FILE_PATH}")
        print("   ⏳ Opening workbook...")
        
        # Tunggu Excel visible di layar
        time.sleep(2)

        start_time = time.time()
        try:
            print(f"   📂 Mencoba buka: {EXCEL_FILE_PATH}")
            print(f"   ⏳ Waiting for Open...")
            
            # Coba dengan parameter minimal - avoid dialog
            wb = excel.Workbooks.Open(
                Filename=EXCEL_FILE_PATH,
                UpdateLinks=False,  # Jangan update external links
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

        # Jeda untuk RefreshAll benar-benar mulai
        time.sleep(2)

        # ✅ TUNGGU SAMPAI QUERY SELESAI - MAX 1800s (30 menit)
        # wait_for_connections_refresh(wb, max_wait=1800, poll_interval=3)
        print("   ✓ Proses refresh GCP selesai!")

        # Parse Group mapping dari sheet Group
        group_mapping = parse_group_mapping(wb)

        print("\n🔄 [4/5] Membaca dan memproses sheet Caption...")

        sheets_to_process = ["Caption", "Caption 2"]

        for sheet_label in sheets_to_process:
            try:
                ws_caption = wb.Sheets(sheet_label)
                print(f"\n{'='*60}")
                print(f"📄 MEMPROSES SHEET: {sheet_label}")
                print(f"{'='*60}")

                capture_config, caption_text = parse_caption_table(ws_caption)

                if not capture_config:
                    print(f"⚠️ Tidak ada dashboard di sheet {sheet_label}, skip")
                    continue

                print(f"\n   ✓ {len(capture_config)} dashboard ditemukan")

                # Group dashboards by consecutive group
                grouped_dashboards = group_dashboards_by_consecutive_group(capture_config)
                print(f"   📊 Dikelompokkan menjadi {len(grouped_dashboards)} group")
                
                for group_idx, group_data in enumerate(grouped_dashboards, 1):
                    group_name = group_data['group']
                    dashboards = group_data['dashboards']
                    
                    print(f"\n{'='*60}")
                    print(f"📸 GROUP {group_idx}/{len(grouped_dashboards)}: {group_name}")
                    print(f"   Dashboards: {len(dashboards)} item")
                    print(f"{'='*60}")
                    
                    # Cari Group IDs dari mapping
                    group_ids = group_mapping.get(group_name, '')
                    if not group_ids:
                        print(f"   ⚠️ Group '{group_name}' tidak ditemukan di sheet Group")
                        continue
                    
                    # Group ID bisa multiple, dipisah dengan ;
                    group_ids_list = [gid.strip() for gid in group_ids.split(';') if gid.strip()]
                    print(f"   📍 Akan dikirim ke: {', '.join(group_ids_list)}")
                    
                    # Screenshot semua dashboard dalam group ini
                    print(f"\n🔄 [5/5] Screenshot untuk group: {group_name}...")
                    time.sleep(1)
                    
                    screenshot_count = 0
                    for dash_idx, config in enumerate(dashboards, 1):
                        try:
                            sheet_name = config['sheet']
                            from_range = config['from']
                            to_range = config['to']
                            dashboard_name = config['dashboard_name']

                            print(f"\n📸 [{dash_idx}/{len(dashboards)}] {dashboard_name}")
                            print(f"   Sheet: {sheet_name} | Range: {from_range}:{to_range}")

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
                                screenshot_count += 1
                                img = img.crop(img.getbbox())

                                enhancer_sharpness = ImageEnhance.Sharpness(img)
                                img = enhancer_sharpness.enhance(1.5)

                                enhancer_contrast = ImageEnhance.Contrast(img)
                                img = enhancer_contrast.enhance(1.2)

                                enhancer_brightness = ImageEnhance.Brightness(img)
                                img = enhancer_brightness.enhance(1.05)

                                image_path = os.path.join(SCRIPT_DIR, f"temp_report_{screenshot_count}.png")

                                img_rgb = img.convert('RGB')
                                img_rgb.save(image_path, 'JPEG', quality=95, optimize=True)
                                print(f"   ✓ Screenshot berhasil: {os.path.basename(image_path)} ({img_rgb.size})")

                                excel.SendKeys('{ESCAPE}')
                            else:
                                print(f"   ⚠️ Clipboard kosong untuk {dashboard_name}")
                                excel.SendKeys('{ESCAPE}')

                        except Exception as e:
                            print(f"   ❌ Error capturing {config['dashboard_name']}: {str(e)[:80]}")
                            excel.SendKeys('{ESCAPE}')
                            continue

                    print(f"\n✅ Group {group_name}: {screenshot_count} screenshot berhasil")

                    if screenshot_count == 0:
                        print(f"⚠️ Tidak ada screenshot untuk group {group_name}")
                        continue

                    # Simpan caption dan group IDs ke file konfigurasi untuk wa_bot
                    save_send_config(caption_text, group_ids_list, screenshot_count)

                    print("\n" + "="*60)
                    print(f"📱 MENGIRIM GROUP: {group_name}")
                    print("="*60)
                    send_wa_report()

                    cleanup_screenshot_files()

                print(f"\n✅ Sheet {sheet_label}: semua group selesai diproses!")

            except Exception as e:
                print(f"⚠️ Error memproses sheet {sheet_label}: {str(e)[:80]}")
                continue

        print("\n" + "="*60)
        print("✅ SEMUA SHEET SELESAI DIPROSES!")
        print("="*60)

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


def save_send_config(caption_text, group_ids, screenshot_count):
    """Simpan konfigurasi untuk wa_bot.js: caption text, group IDs, dan jumlah screenshots"""
    import json
    
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


def cleanup_screenshot_files():
    """Hapus semua file screenshot temporary setelah dikirim"""
    try:
        print("\n🗑️ Membersihkan file screenshot temporary...")

        screenshot_files = glob.glob(os.path.join(SCRIPT_DIR, "temp_report_*.png"))
        for file in screenshot_files:
            os.remove(file)
            print(f"   ✓ Dihapus: {os.path.basename(file)}")

        send_config_file = os.path.join(SCRIPT_DIR, "send_config.json")
        if os.path.exists(send_config_file):
            os.remove(send_config_file)
            print(f"   ✓ Dihapus: send_config.json")

        print("✅ Cleanup selesai!")

    except Exception as e:
        print(f"⚠️ Error saat cleanup: {str(e)}")


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