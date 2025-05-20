import pandas as pd

# Baca semua sheet
data = pd.read_excel('author79.xlsx', sheet_name=['Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'])

# Ambil daftar nama dari Sheet5
nama_list = data['Sheet5']['Nama'].dropna().unique()


# Fungsi untuk mencari fakultas dan nama asli dari sumber
def cari_fakultas(nama):
    hasil_pencarian = []

    # Cek di Sheet2
    if 'Nama' in data['Sheet2'].columns and 'Fakultas' in data['Sheet2'].columns:
        result = data['Sheet2'][data['Sheet2']['Nama'].str.contains(nama, case=False, na=False)]
        for _, row in result.iterrows():
            hasil_pencarian.append({
                'Nama_Asli': row['Nama'],
                'Fakultas': row['Fakultas'],
                'Sumber': 'Sheet2'
            })

    # Cek di Sheet3
    if 'Nama' in data['Sheet3'].columns and 'Fakultas' in data['Sheet3'].columns:
        result = data['Sheet3'][data['Sheet3']['Nama'].str.contains(nama, case=False, na=False)]
        for _, row in result.iterrows():
            hasil_pencarian.append({
                'Nama_Asli': row['Nama'],
                'Fakultas': row['Fakultas'],
                'Sumber': 'Sheet3'
            })

    # Cek di Sheet4
    if 'Nama' in data['Sheet4'].columns and 'Fakultas' in data['Sheet4'].columns:
        result = data['Sheet4'][data['Sheet4']['Nama'].str.contains(nama, case=False, na=False)]
        for _, row in result.iterrows():
            hasil_pencarian.append({
                'Nama_Asli': row['Nama'],
                'Fakultas': row['Fakultas'],
                'Sumber': 'Sheet4'
            })

    return hasil_pencarian if hasil_pencarian else [{
        'Nama_Asli': nama,
        'Fakultas': 'Tidak Ditemukan',
        'Sumber': ''
    }]


# Buat list untuk menampung hasil
hasil_list = []

for nama in nama_list:
    # Cari data fakultas
    pencarian = cari_fakultas(nama)

    for hasil in pencarian:
        hasil_list.append({
            'Nama_Pencarian': nama,
            'Nama_Asli_Di_Sumber': hasil['Nama_Asli'],
            'Fakultas': hasil['Fakultas'],
            'Sumber_Sheet': hasil['Sumber']
        })

# Konversi list ke DataFrame
hasil = pd.DataFrame(hasil_list)

# Simpan ke Excel
hasil.to_excel('hasil_pencarian_fakultas.xlsx', index=False)

print("Proses selesai. Hasil disimpan di 'hasil_pencarian_fakultas.xlsx'")
print("\nContoh hasil:")
print(hasil.head(10).to_string(index=False))