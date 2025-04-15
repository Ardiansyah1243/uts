import csv
import os
from datetime import date
import openpyxl

# File yang digunakan
file_csv = "data_zakat.csv"
file_excel = "report_zakat.xlsx"
file_harga_beras = "harga_beras.csv"

# Load harga beras dari file
def load_harga_beras():
    harga = []
    if os.path.exists(file_harga_beras):
        with open(file_harga_beras, newline="") as file:
            reader = csv.reader(file)
            for row in reader:
                if row:
                    harga.append(int(row[0]))
    return harga

# Simpan harga beras ke file
def simpan_harga_beras(harga):
    with open(file_harga_beras, mode="w", newline="") as file:
        writer = csv.writer(file)
        for h in harga:
            writer.writerow([h])

# Tampilkan harga beras
def tampilkan_harga_beras(harga_list):
    if not harga_list:
        print("Belum ada harga beras.")
    for i, h in enumerate(harga_list):
        print(f"({i+1}) Rp {h}")

# Tambah harga beras
def input_harga_beras(harga_list):
    try:
        harga = int(input("Masukkan Harga Beras Per-Kilo: "))
        harga_list.append(harga)
        simpan_harga_beras(harga_list)
    except ValueError:
        print("Input harus angka.")

# Hapus semua harga beras
def hapus_harga_beras():
    if os.path.exists(file_harga_beras):
        os.remove(file_harga_beras)
        print("Semua data harga beras telah dihapus.")
    else:
        print("Tidak ada data harga beras yang perlu dihapus.")

# Simpan data zakat ke file CSV
def simpan_data_csv(data):
    file_exists = os.path.exists(file_csv)
    with open(file_csv, mode="a", newline="") as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(["NIK", "Nama", "Tanggal Bayar", "Harga Per-Kg", "Jumlah Kepala"])
        for row in data:
            writer.writerow(row)

# Tampilkan data zakat dari CSV
def tampilkan_data():
    if not os.path.exists(file_csv):
        print("Belum ada data zakat.")
        return
    with open(file_csv, newline="") as file:
        reader = csv.reader(file)
        for row in reader:
            print(row)

# Hapus semua data zakat
def hapus_data_zakat():
    if os.path.exists(file_csv):
        os.remove(file_csv)
        print("Semua data zakat berhasil dihapus.")
    else:
        print("Belum ada data zakat yang perlu dihapus.")

# Proses pembayaran zakat
def pembayaran_zakat(harga_list):
    tampilkan_harga_beras(harga_list)
    try:
        nik = input("Masukkan NIK: ")
        nama = input("Masukkan Nama: ")
        id_beras = int(input("Pilih nomor harga beras: "))
        jumlah_kepala = int(input("Masukkan Jumlah Kepala: "))
        harga_total = harga_list[id_beras - 1] * jumlah_kepala
        print("Total yang harus dibayar: Rp", harga_total)
        bayar = int(input("Masukkan jumlah uang yang dibayar: "))
        kembali = bayar - harga_total
        print("Kembalian Anda: Rp", kembali)

        data = [(nik, nama, date.today(), harga_list[id_beras - 1], jumlah_kepala)]
        simpan_data_csv(data)
    except (ValueError, IndexError):
        print("Input tidak valid. Pastikan semua angka dimasukkan dengan benar.")

# Export data ke Excel
def export_to_excel():
    if not os.path.exists(file_csv):
        print("Tidak ada data untuk di-export.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Zakat"

    with open(file_csv, newline="") as file:
        reader = csv.reader(file)
        for row in reader:
            ws.append(row)

    wb.save(file_excel)
    print(f"Data berhasil diekspor ke {file_excel}")

# Main Program
def main():
    harga_beras = load_harga_beras()

    while True:
        print("""
Menu:
1. Tampilkan Harga Beras
2. Input Harga Beras
3. Hapus Semua Harga Beras
4. Tampilkan Data Zakat
5. Pembayaran Zakat
6. Hapus Semua Data Zakat
7. Export ke Excel
8. Keluar
""")
        pilihan = input("Pilih menu (1-8): ")

        if pilihan == '1':
            tampilkan_harga_beras(harga_beras)
        elif pilihan == '2':
            input_harga_beras(harga_beras)
        elif pilihan == '3':
            hapus_harga_beras()
            harga_beras = []
        elif pilihan == '4':
            tampilkan_data()
        elif pilihan == '5':
            if not harga_beras:
                print("Masukkan dulu harga beras (menu 2).")
            else:
                pembayaran_zakat(harga_beras)
        elif pilihan == '6':
            hapus_data_zakat()
        elif pilihan == '7':
            export_to_excel()
        elif pilihan == '8':
            print("Keluar dari program.")
            break
        else:
            print("Pilihan tidak valid.")

if __name__ == "__main__":
    main()