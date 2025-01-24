import subprocess
import os


def download_file():
  
  tahun = "2023"
  kuartal_romawi = "II"
  kuartal = "TW2"
  kode_perusahaan = "ACES"
  download_dir = f"data_ori_{tahun}_{kuartal}"

  # Membuat direktori jika belum ada
  if not os.path.exists(download_dir):
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

if __name__ == "__main__":
  download_file()


