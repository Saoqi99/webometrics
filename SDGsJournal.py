import pandas as pd
import time
from scholarly import scholarly

# Kata kunci SDGs (Bahasa Inggris & Indonesia)
sdgs_keywords = [
    # English
    'sustainable development', 'climate change', 'poverty', 'gender equality',
    'clean energy', 'quality education', 'zero hunger', 'life on land',
    'life below water', 'partnerships', 'inequality', 'green economy',
    'sustainable cities', 'responsible consumption', 'clean water',
    # Indonesian
    'pembangunan berkelanjutan', 'perubahan iklim', 'kemiskinan', 'kesetaraan gender',
    'energi bersih', 'pendidikan berkualitas', 'tanpa kelaparan', 'kehidupan di darat',
    'kehidupan bawah laut', 'kemitraan', 'pengurangan ketimpangan',
    'ekonomi hijau', 'kota berkelanjutan', 'konsumsi bertanggung jawab', 'air bersih',
    'sanitasi layak', 'pertumbuhan ekonomi', 'iklim', 'lingkungan', 'keadilan sosial'
]

# Cek apakah judul terkait SDGs
def is_sdgs_related(title, keywords=sdgs_keywords):
    return any(keyword in title.lower() for keyword in keywords)

# Fungsi utama
def find_sdgs_by_id(csv_path="uin_authors.csv", year_filter=2024, delay=15):
    try:
        df_authors = pd.read_csv(csv_path)
    except Exception as e:
        print(f"❌ Gagal membaca file CSV: {e}")
        return

    results = []

    for i, row in df_authors.iterrows():
        scholar_id = row.get("scholar_id", "").strip()
        name = row.get("name", "").strip()

        print(f"\n🔍 Mengambil data dari: {name} (ID: {scholar_id})")

        try:
            author = scholarly.search_author_id(scholar_id)
            if not author:
                print(f"⚠️ Tidak ditemukan profil dengan ID: {scholar_id}")
                continue

            filled_author = scholarly.fill(author)
            if not filled_author or not isinstance(filled_author, dict):
                print(f"⚠️ Profil tidak bisa dimuat dengan benar untuk ID: {scholar_id}")
                continue

            affil = filled_author.get("affiliation", "Tidak tersedia")

            for pub in filled_author.get("publications", []):
                time.sleep(delay)

                try:
                    pub_details = scholarly.fill(pub)
                    if not pub_details or "bib" not in pub_details:
                        continue

                    title = pub_details["bib"].get("title", "")
                    year = pub_details["bib"].get("pub_year")

                    if year and int(year) == year_filter and is_sdgs_related(title):
                        results.append({
                            "Author": name,
                            "Title": title,
                            "Year": year,
                            "Citations": pub_details.get("num_citations", 0),
                            "Affiliation": affil,
                            "Scholar URL": f"https://scholar.google.com/scholar?oi=bibs&hl=en&q={title.replace(' ', '+')}"
                        })

                except Exception as e:
                    print(f"⚠️ Gagal membaca publikasi: {e}")
                    continue

        except Exception as e:
            print(f"❌ Error mengambil profil: {e}")
            continue

    # Simpan hasil ke Excel
    df_result = pd.DataFrame(results)
    filename = f"sdgs_2025_by_id.xlsx"
    df_result.to_excel(filename, index=False)
    print(f"\n✅ {len(df_result)} publikasi SDGs ditemukan dan disimpan ke: {filename}")
    return df_result

# Eksekusi
if __name__ == "__main__":
    find_sdgs_by_id(csv_path="uin_authors.csv", year_filter=2024, delay=15)
