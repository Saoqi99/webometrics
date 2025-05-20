import pandas as pd

# Baca data
df = pd.read_excel('author79.xlsx', sheet_name='Sheet1')

# 1. Proses data nama (N1, N2, N3, AllN)
nama_data = pd.DataFrame()
for col in ['N1', 'N2', 'N3', 'AllN']:
    temp = df[[col]].dropna().rename(columns={col: 'Nama'})
    temp['Kolom_Asal'] = col
    temp['Tipe'] = 'Nama'
    nama_data = pd.concat([nama_data, temp], ignore_index=True)

# 2. Proses data ID (S1, S2, S3, AllS)
id_data = pd.DataFrame()
for col in ['S1', 'S2', 'S3', 'AllS']:
    temp = df[[col]].dropna().rename(columns={col: 'ID'})
    temp['Kolom_Asal'] = col
    temp['Tipe'] = 'ID'
    id_data = pd.concat([id_data, temp], ignore_index=True)

# 3. Hapus duplikat secara terpisah
nama_unik = nama_data.drop_duplicates(subset=['Nama'], keep='first')
id_unik = id_data.drop_duplicates(subset=['ID'], keep='first')

# 4. Tampilkan hasil
print("DAFTAR NAMA UNIK")
print("="*50)
print(nama_unik.sort_values('Nama').to_string(index=False))

print("\nDAFTAR ID UNIK")
print("="*50)
print(id_unik.sort_values('ID').to_string(index=False))

# 5. Simpan ke Excel
with pd.ExcelWriter('author79_clean.xlsx') as writer:
    nama_unik.to_excel(writer, sheet_name='Nama_Unik', index=False)
    id_unik.to_excel(writer, sheet_name='ID_Unik', index=False)

# 6. Statistik
print("\nSTATISTIK NAMA:")
print(nama_unik['Kolom_Asal'].value_counts())
print("\nSTATISTIK ID:")
print(id_unik['Kolom_Asal'].value_counts())