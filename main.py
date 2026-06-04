import os
import sys
import time
import subprocess
from datetime import datetime
import win32com.client as win32
from PIL import ImageGrab

# Gunakan absolute path berdasarkan lokasi file ini
WORK_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(WORK_DIR, "dashboard_sales.xlsx")
IMAGE_SAVE_PATH = os.path.join(WORK_DIR, "temp_report.png")
SHEET_NAME = "Dashboard"

def cek_jadwal():
    hari_ini = datetime.now().day
    if hari_ini >= 20:
        print(f"Tanggal {hari_ini}. Jadwal kirim tiap hari berjalan.")
        return True
    elif hari_ini % 2 == 0: 
        print(f"Tanggal {hari_ini}. Jadwal kirim 2 hari sekali berjalan.")
        return True
    else:
        print(f"Tanggal {hari_ini}. Bukan jadwal pengiriman.")
        return False

def refresh_dan_screenshot():
    print("Membuka Excel dan menyinkronkan data GCP...")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone() # Tunggu data GCP selesai ditarik
        print("Data GCP berhasil ditarik!")
        
        # Iterasi semua sheet yang namanya mengandung "Dashboard"
        # Iterasi semua sheet yang namanya mengandung "Dashboard"
        # Prioritas: "Dashboard" (exact) dulu, terus "Dashboard 1", "Dashboard 2", dst
        dashboard_sheets = []
        for ws in wb.Sheets:
            if "Dashboard" in ws.Name:
                dashboard_sheets.append(ws)
        
        # Sort: "Dashboard" terlebih dahulu, terus yang lainnya
        dashboard_sheets.sort(key=lambda x: (x.Name != "Dashboard", x.Name))
        
        dashboard_count = 0
        for ws in dashboard_sheets:
            print(f"Memproses sheet: {ws.Name}")
            
            try:
                # Cari baris dan kolom terakhir yang ada data
                last_row = ws.Cells.SpecialCells(11).Row  # 11 = xlCellTypeLastCell
                last_col = ws.Cells.SpecialCells(11).Column
                
                # Validasi: minimal harus ada data (lebih dari 1x1)
                if last_row < 1 or last_col < 1:
                    print(f"  ⚠️ Sheet {ws.Name} tidak punya data yang valid. Skip.")
                    continue
                
                # Activate dan select range
                ws.Activate()
                time.sleep(0.3)
                
                # Tentukan range dari A1 sampai ke sel terakhir
                tabel_range = ws.Range(f"A1:{ws.Cells(last_row, last_col).Address}")
                tabel_range.Select()
                time.sleep(0.3)
                
                # Coba copy picture dengan Copy biasa dulu
                try:
                    # Method 1: Gunakan Copy untuk clipboard
                    ws.Range(f"A1:{ws.Cells(last_row, last_col).Address}").Copy()
                    time.sleep(0.5)
                    
                    # Ambil dari clipboard
                    img = ImageGrab.grabclipboard()
                    if img:
                        dashboard_count += 1
                        # Auto-crop gambar
                        img = img.crop(img.getbbox())
                        image_path = os.path.join(WORK_DIR, f"temp_report_{dashboard_count}.png")
                        img.save(image_path, 'PNG')
                        print(f"  ✓ Screenshot {ws.Name} berhasil disimpan. Ukuran: {img.size}")
                        # Escape dari copy mode
                        excel.SendKeys('{ESCAPE}')
                    else:
                        print(f"  ⚠️ Clipboard kosong untuk {ws.Name}.")
                        excel.SendKeys('{ESCAPE}')
                        
                except Exception as copy_error:
                    print(f"  ⚠️ Gagal copy dari {ws.Name}: {str(copy_error)[:80]}")
                    excel.SendKeys('{ESCAPE}')
                    continue
                    
            except Exception as e:
                print(f"  ⚠️ Skipping sheet {ws.Name}: {str(e)[:80]}")
                continue
        
        if dashboard_count == 0:
            print("Tidak ditemukan sheet dengan nama Dashboard.")
            sys.exit(1)
        
        # Ambil caption dari sheet "Caption" mulai dari A2
        try:
            ws_caption = wb.Sheets("Caption")
            caption_lines = []
            
            row = 2
            while True:
                cell_value = ws_caption.Cells(row, 1).Value
                if cell_value is None:
                    break
                caption_lines.append(str(cell_value).strip())
                row += 1
            
            if caption_lines:
                caption_text = "\n".join(caption_lines)
                # Simpan caption ke file
                caption_path = os.path.join(WORK_DIR, "caption.txt")
                with open(caption_path, "w", encoding="utf-8") as f:
                    f.write(caption_text)
                print(f"Caption berhasil diambil ({len(caption_lines)} baris).")
            else:
                print("Sheet Caption kosong.")
        except Exception as e:
            print(f"Sheet Caption tidak ditemukan atau error: {e}")
            
    finally:
        wb.Save()
        wb.Close()
        excel.Quit()

def trigger_baileys():
    print("Menjalankan WhatsApp Bot (Baileys)...")
    # Memanggil script Node.js dari Python
    wa_bot_path = os.path.join(WORK_DIR, "wa_bot.js")
    proses = subprocess.run(["node", wa_bot_path], capture_output=False, text=True, cwd=WORK_DIR)
    
    if proses.returncode == 0:
        print("Siklus automasi selesai dengan sukses.")
        
        # Hapus semua file temp_report_*.png dan caption.txt setelah berhasil dikirim
        import glob
        for file in glob.glob(os.path.join(WORK_DIR, "temp_report_*.png")):
            try:
                os.remove(file)
                print(f"File {os.path.basename(file)} dihapus.")
            except Exception as e:
                print(f"Gagal menghapus {file}: {e}")
        
        caption_path = os.path.join(WORK_DIR, "caption.txt")
        if os.path.exists(caption_path):
            try:
                os.remove(caption_path)
                print("File caption.txt dihapus.")
            except Exception as e:
                print(f"Gagal menghapus caption.txt: {e}")
    else:
        print("Terjadi kendala pada pengiriman WhatsApp.")

if __name__ == "__main__":
    if cek_jadwal():
        refresh_dan_screenshot()
        # Cek apakah ada file screenshot yang dibuat
        import glob
        if glob.glob(os.path.join(WORK_DIR, "temp_report_*.png")):
            trigger_baileys()