import os
import sys
import time
import subprocess
from datetime import datetime
import win32com.client as win32
from PIL import ImageGrab
import glob

# Path untuk file reports (folder File report)
WORK_DIR = r"C:\Users\fresn\OneDrive - Indosatooredoo Hutchison\File report"

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

def parse_caption_table(ws_caption):
    """Parse sheet Caption:
       1. Cari header (Dashboard Name, Sheet, From, To) - bisa di row mana saja
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
        
        if 'sheet' in headers:
            sheet_name = ws_caption.Cells(row, headers['sheet']).Value
        if 'from' in headers:
            from_range = ws_caption.Cells(row, headers['from']).Value
        if 'to' in headers:
            to_range = ws_caption.Cells(row, headers['to']).Value
        
        if sheet_name and from_range and to_range:
            data_list.append({
                'dashboard_name': str(dashboard_name).strip(),
                'sheet': str(sheet_name).strip(),
                'from': str(from_range).strip(),
                'to': str(to_range).strip(),
            })
            print(f"   ✓ Row {row}: {dashboard_name} | Sheet: {sheet_name} | Range: {from_range}:{to_range}")
        
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

def refresh_dan_screenshot():
    print("\n" + "="*60)
    print("📊 MEMBUKA FILE DAN REFRESH DATA GCP")
    print("="*60)
    
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False  # Jalankan di background, tidak tampil di layar
    excel.DisplayAlerts = False
    wb = None  # Initialize wb agar tidak error di finally block
    
    try:
        print(f"\n🔄 [1/5] Membuka file: {os.path.basename(EXCEL_FILE_PATH)}")
        print(f"    Path: {EXCEL_FILE_PATH}")
        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        print("   ✓ File terbuka di background")
        
        print("🔄 [2/5] Melakukan refresh data dari GCP...")
        wb.RefreshAll()
        
        print("🔄 [3/5] Menunggu query GCP selesai...")
        excel.CalculateUntilAsyncQueriesDone()
        print("   ✓ Data GCP berhasil di-refresh!")
        
        print("🔄 [4/5] Membaca dan memproses sheet Caption...")
        
        # Loop per sheet: Caption, Caption 2, dst
        sheets_to_process = ["Caption", "Caption 2"]
        
        for sheet_label in sheets_to_process:
            try:
                ws_caption = wb.Sheets(sheet_label)
                print(f"\n{'='*60}")
                print(f"📄 MEMPROSES SHEET: {sheet_label}")
                print(f"{'='*60}")
                
                # Baca config dan caption dari sheet
                capture_config, caption_text = parse_caption_table(ws_caption)
                
                if not capture_config:
                    print(f"⚠️ Tidak ada dashboard di sheet {sheet_label}, skip")
                    continue
                
                print(f"\n   ✓ {len(capture_config)} dashboard ditemukan")
                
                # Screenshot semua dashboard di sheet ini
                print("\n🔄 [5/5] Memulai screenshot dashboard...")
                time.sleep(2)
                
                screenshot_count = 0
                for i, config in enumerate(capture_config, 1):
                    try:
                        sheet_name = config['sheet']
                        from_range = config['from']
                        to_range = config['to']
                        dashboard_name = config['dashboard_name']
                        
                        print(f"\n📸 [{i}/{len(capture_config)}] Capturing: {dashboard_name}")
                        print(f"   Sheet: {sheet_name} | Range: {from_range}:{to_range}")
                        
                        # Aktifkan sheet
                        ws = wb.Sheets(sheet_name)
                        ws.Activate()
                        time.sleep(0.5)
                        
                        # Select range
                        range_address = f"{from_range}:{to_range}"
                        selected_range = ws.Range(range_address)
                        selected_range.Select()
                        time.sleep(0.3)
                        
                        # Copy range
                        selected_range.Copy()
                        time.sleep(0.5)
                        
                        # Grab dari clipboard
                        img = ImageGrab.grabclipboard()
                        if img:
                            screenshot_count += 1
                            img = img.crop(img.getbbox())
                            
                            # Buat nama file sesuai format wa_bot.js: temp_report_1.png, temp_report_2.png, dst
                            # Simpan di SCRIPT_DIR agar wa_bot.js bisa menemukannya
                            image_path = os.path.join(SCRIPT_DIR, f"temp_report_{screenshot_count}.png")
                            img.save(image_path, 'PNG')
                            print(f"   ✓ Screenshot berhasil: {os.path.basename(image_path)} ({img.size})")
                            
                            excel.SendKeys('{ESCAPE}')
                        else:
                            print(f"   ⚠️ Clipboard kosong untuk {dashboard_name}")
                            excel.SendKeys('{ESCAPE}')
                            
                    except Exception as e:
                        print(f"   ❌ Error capturing {config['dashboard_name']}: {str(e)[:80]}")
                        excel.SendKeys('{ESCAPE}')
                        continue
                
                print(f"\n✅ Sheet {sheet_label}: {screenshot_count} screenshot berhasil")
                
                if screenshot_count == 0:
                    print(f"⚠️ Tidak ada screenshot yang berhasil untuk {sheet_label}")
                    continue
                
                # Simpan caption untuk sheet ini
                save_caption_to_file(caption_text)
                
                # Kirim ke WA untuk sheet ini
                print("\n" + "="*60)
                print(f"📱 MENGIRIM {sheet_label} VIA WHATSAPP")
                print("="*60)
                send_wa_report()
                
                # Cleanup screenshot untuk sheet ini
                cleanup_screenshot_files()
                
            except Exception as e:
                print(f"⚠️ Error memproses sheet {sheet_label}: {str(e)[:80]}")
                continue
        
        print("\n" + "="*60)
        print("✅ SEMUA SHEET SELESAI DIPROSES!")
        print("="*60)
        
    finally:
        try:
            if wb:  # Check apakah wb sudah berhasil dibuka
                wb.Save()
        except:
            pass  # File mungkin read-only, ignore error
        
        if wb:  # Check sebelum close
            wb.Close()
        
        excel.Quit()

def send_wa_report():
    """Jalankan wa_bot.js untuk mengirim screenshot ke WA"""
    try:
        print("\n🔄 Menjalankan WA Bot untuk mengirim screenshot...")
        
        # Jalankan wa_bot.js menggunakan Node.js dari SCRIPT_DIR (tempat wa_bot.js berada)
        result = subprocess.run(
            ["node", os.path.join(SCRIPT_DIR, "wa_bot.js")],
            cwd=SCRIPT_DIR,
            capture_output=True,
            text=True,
            timeout=300  # Timeout 5 menit
        )
        
        if result.returncode == 0:
            print("✅ WA Bot selesai!")
            if result.stdout:
                print(result.stdout)
         
            # Hapus file screenshot setelah berhasil dikirim
            cleanup_screenshot_files()
        else:
            print("❌ WA Bot gagal dengan error:")
            if result.stderr:
                print(result.stderr)
            if result.stdout:
                print(result.stdout)
            raise RuntimeError("WA Bot gagal mengirim report")
            
    except subprocess.TimeoutExpired:
        print("❌ WA Bot timeout (terlalu lama)")
        raise
    except Exception as e:
        print(f"❌ Error menjalankan WA Bot: {str(e)}")
        raise

def save_caption_to_file(caption_text):
    """Simpan caption text ke file caption.txt"""
    try:
        # Simpan di SCRIPT_DIR agar wa_bot.js bisa menemukannya
        caption_file = os.path.join(SCRIPT_DIR, "caption.txt")
        
        if caption_text and caption_text.strip():
            with open(caption_file, 'w', encoding='utf-8') as f:
                f.write(caption_text)
            print(f"\n📝 Caption disimpan ke file: caption.txt")
            print(f"   Isi:\n{caption_text}")
        else:
            # Jika tidak ada caption, buat file kosong
            with open(caption_file, 'w', encoding='utf-8') as f:
                f.write("")
            print(f"\n📝 File caption.txt dibuat (kosong)")
        
    except Exception as e:
        print(f"⚠️ Error menyimpan caption: {str(e)}")

def cleanup_screenshot_files():
    """Hapus semua file screenshot temporary setelah dikirim"""
    try:
        print("\n🗑️ Membersihkan file screenshot temporary...")
        
        # Hapus file temp_report_*.png dari SCRIPT_DIR
        screenshot_files = glob.glob(os.path.join(SCRIPT_DIR, "temp_report_*.png"))
        for file in screenshot_files:
            os.remove(file)
            print(f"   ✓ Dihapus: {os.path.basename(file)}")
        
        # Hapus file caption.txt jika ada
        caption_file = os.path.join(SCRIPT_DIR, "caption.txt")
        if os.path.exists(caption_file):
            os.remove(caption_file)
            print(f"   ✓ Dihapus: caption.txt")
        
        print("✅ Cleanup selesai!")
        
    except Exception as e:
        print(f"⚠️ Error saat cleanup: {str(e)}")

if __name__ == "__main__":
    refresh_dan_screenshot()
