saya ingin agar programnya melakukan generate link secara otomatisa dengan perulangan dari data yang saya berikan
pertama, tahun saya ingin agar mendownload dari tahun 2020 sampai 2024

kedua, kuartal ada "TW1", "TW2", "TW3", dan "Audit"

ketiga, untuk kuartal_romawinya itu ada "I" untuk kuartal "TW1", "II" untuk kuartal "TW2", "III" untuk kuartal "TW3", dan "Tahunan" untuk kuartal "Audit"

untuk kode perusahaan itu saya simpan di file excel dengan nama "database_nama_perusahaan", sheet1, kolom A

jadi nanti yang pertama kita tentukan tahunnya dulu, lalu kuartalnya, lalu kode perusahaannya

dan ini program yang saya buat sebelumnya


import subprocess

import os

defdownload_file():

  tahun = "2023"

  kuartal_romawi = "II"

  kuartal = "TW2"

  kode_perusahaan = "ACES"

  download_dir = f"data_ori_{tahun}_{kuartal}"

# Membuat direktori jika belum ada

ifnot os.path.exists(download_dir):

    os.makedirs(download_dir)

# Path lengkap file output

  output_file = os.path.join(download_dir, f"FinancialStatement-{tahun}-{kuartal_romawi}-{kode_perusahaan}.xlsx")

  command = [

"curl",

"-H", "User-Agent: Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/110.0",

f"https://idx.co.id/Portals/0/StaticData/ListedCompanies/Corporate_Actions/New_Info_JSX/Jenis_Informasi/01_Laporan_Keuangan/02_Soft_Copy_Laporan_Keuangan//Laporan%20Keuangan%20Tahun%20{tahun}/{kuartal}/{kode_perusahaan}/FinancialStatement-{tahun}-{kuartal_romawi}-{kode_perusahaan}.xlsx",

"-o", output_file

  ]

try:

    subprocess.run(command, check=True)

print(f"File berhasil diunduh ke {output_file}")

except subprocess.CalledProcessError as e:

print(f"Error: {e}")

if __name__ =="__main__":

download_file()
