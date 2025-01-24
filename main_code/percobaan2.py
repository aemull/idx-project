import subprocess
import os
import pandas as pd
from datetime import datetime

def log_error(message):
    """Mencatat pesan kesalahan ke dalam file log."""
    with open("download_errors.log", "a") as log_file:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"[{timestamp}] {message}\n")

def download_files():
    # Rentang tahun
    tahun_range = range(2022, 2025)
    
    # Kuartal dan Romawi
    kuartal_dict = {
        "TW1": "I",
        "TW2": "II",
        "TW3": "III",
        "Audit": "Tahunan"
    }
    
    # Membaca kode perusahaan dari file Excel
    excel_file = "contoh_format_data/database_nama_perusahaan.xlsx"
    sheet_name = "Sheet1"
    
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        kode_perusahaan_list = df['Kode Perusahaan'].dropna().tolist()  # Pastikan kolom bernama "Kode Perusahaan"
    except Exception as e:
        log_error(f"Error membaca file Excel: {e}")
        print(f"Error membaca file Excel: {e}")
        return

    # Iterasi tahun, kuartal, dan kode perusahaan
    for tahun in tahun_range:
        for kuartal, kuartal_romawi in kuartal_dict.items():
            for kode_perusahaan in kode_perusahaan_list:
                # Membuat direktori berdasarkan tahun dan kuartal
                download_dir = f"data_ori_{tahun}_{kuartal}"
                if not os.path.exists(download_dir):
                    os.makedirs(download_dir)

                # Path lengkap file output
                output_file = os.path.join(download_dir, f"FinancialStatement-{tahun}-{kuartal_romawi}-{kode_perusahaan}.xlsx")
                
                # URL file yang akan diunduh
                url = (
                    f"https://idx.co.id/Portals/0/StaticData/ListedCompanies/Corporate_Actions/New_Info_JSX/Jenis_Informasi/01_Laporan_Keuangan/"
                    f"02_Soft_Copy_Laporan_Keuangan/Laporan%20Keuangan%20Tahun%20{tahun}/{kuartal}/{kode_perusahaan}/"
                    f"FinancialStatement-{tahun}-{kuartal_romawi}-{kode_perusahaan}.xlsx"
                )
                
                # Perintah curl
                command = [
                    "curl",
                    "-H", "User-Agent: Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/110.0",
                    url,
                    "-o", output_file
                ]
                
                # Eksekusi perintah
                try:
                    subprocess.run(command, check=True)
                    print(f"File berhasil diunduh: {output_file}")
                except subprocess.CalledProcessError as e:
                    error_message = f"Gagal mengunduh file: {url}. Error: {e}"
                    log_error(error_message)
                    print(error_message)

if __name__ == "__main__":
    download_files()
