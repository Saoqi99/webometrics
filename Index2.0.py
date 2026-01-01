import time
import random
import pandas as pd
import re
from scholarly import scholarly
from rapidfuzz import fuzz
from scholarly import ProxyGenerator

# Konfigurasi
TARGET_AFFILIATIONS = [
    "Universitas Islam Negeri Sunan Kalijaga",
]
AFFILIATION_THRESHOLD = 75  # Tingkat kemiripan minimal
AUTHOR_MATCH_THRESHOLD = 80  # Kemiripan nama penulis
MAX_CANDIDATES = 5
MAX_PUBLICATIONS = 100
REQUEST_DELAY = (1, 5)  # Jeda antar request


def fetch_candidate_profiles(query, max_results=10):
    """Ambil profil kandidat berdasarkan query pencarian"""
    print(f"\n🔍 Mencari dengan query: '{query}'")
    candidates = []
    try:
        search_results = scholarly.search_author(query)
        for i, author in enumerate(search_results):
            if i >= max_results:
                break
            try:
                author = scholarly.fill(author, sections=['basics'])
                candidates.append({
                    'scholar_id': author.get('scholar_id'),
                    'name': author.get('name'),
                    'affiliation': author.get('affiliation', '').lower(),
                    'email': author.get('email', '').lower()
                })
                print(f"  ✅ Kandidat {i + 1}: {author.get('name')}")
                time.sleep(random.uniform(*REQUEST_DELAY))
            except Exception as e:
                print(f"    🚨 Gagal memproses kandidat {i + 1}: {str(e)}")
    except Exception as e:
        print(f"🚨 Error saat mencari dengan query '{query}': {str(e)}")
    return candidates


def filter_profiles(candidates):
    """Filter kandidat berdasarkan afiliasi dan email"""
    valid_profiles = []
    for profile in candidates:
        try:
            aff_scores = [fuzz.token_sort_ratio(profile['affiliation'], ta.lower())
                          for ta in TARGET_AFFILIATIONS]
            if max(aff_scores) < AFFILIATION_THRESHOLD:
                print(f"⚠️  Afiliasi tidak cocok: {profile['name']} ({profile['affiliation']})")
                continue

            if 'student.uin-suka.ac.id' in profile['email']:
                print(f"⚠️  Email student: {profile['name']}")
                continue

            valid_profiles.append(profile)
        except:
            continue
    return valid_profiles


def parse_authors(authors_str):
    """Parse string penulis ke dalam list nama"""
    if not authors_str:
        return []
    return re.split(r', | and | & ', authors_str.strip())


def check_author_match(owner_name, authors_list):
    """Cek apakah owner_name ada di list penulis"""
    for author in authors_list:
        if fuzz.token_set_ratio(owner_name.lower(), author.lower()) >= AUTHOR_MATCH_THRESHOLD:
            return True
    return False


def process_profile(scholar_id):
    """Proses satu profil dan publikasinya dengan deteksi publikasi meragukan"""
    print(f"\n📖 Memproses profil {scholar_id}")
    try:
        author = scholarly.search_author_id(scholar_id)
        author = scholarly.fill(author, sections=['publications'])
    except Exception as e:
        print(f"   🚨 Gagal memuat profil: {str(e)}")
        return None

    results = []
    owner_name = author.get('name', '')
    pub_count = min(len(author['publications']), MAX_PUBLICATIONS)

    print(f"   🔎 Memeriksa {pub_count} publikasi...")
    for i, pub in enumerate(author['publications'][:pub_count]):
        try:
            publication = scholarly.fill(pub)
            time.sleep(random.uniform(*REQUEST_DELAY))

            bib = publication.get('bib', {})
            authors = parse_authors(bib.get('author', ''))

            # Perbaikan: Handle tahun yang kosong/tidak valid
            year = bib.get('year', '')
            if year:
                try:
                    year = int(year)
                except (ValueError, TypeError):
                    year = str(year).strip()
            else:
                year = 'Tidak Diketahui'

            # Filter hanya tahun 2022–2025
            if isinstance(year, int) and (year < 2022 or year > 2025):
                continue

            nama_cocok = check_author_match(owner_name, authors)
            afiliasi_cocok = any(fuzz.partial_ratio(affil.lower(), bib.get('journal', '').lower()) > 75
                                 for affil in TARGET_AFFILIATIONS)

            results.append({
                'title': bib.get('title', 'No Title'),
                'authors': ', '.join(authors),
                'year': year,
                'journal': bib.get('journal', ''),
                'nama_cocok': 'YA' if nama_cocok else 'TIDAK',
                'afiliasi_cocok': 'YA' if afiliasi_cocok else 'TIDAK',
                'status': 'MERAGUKAN' if not nama_cocok or not afiliasi_cocok else 'VALID'
            })
        except Exception as e:
            print(f"    🚨 Gagal memproses publikasi {i + 1}: {str(e)}")

    return {
        'scholar_id': scholar_id,
        'nama': owner_name,
        'afiliasi': author.get('affiliation', ''),
        'email': author.get('email', ''),
        'publikasi': results
    }


def start():
    """Fungsi utama untuk menjalankan proses"""

    # Langkah 1: Ambil semua kandidat
    all_candidates = []
    for affiliation in TARGET_AFFILIATIONS:
        candidates = fetch_candidate_profiles(affiliation, MAX_CANDIDATES)
        all_candidates.extend(candidates)

    # Langkah 2: Filter kandidat
    valid_profiles = filter_profiles(all_candidates)
    scholar_ids = list({p['scholar_id'] for p in valid_profiles})
    print(f"\n🔍 Ditemukan {len(scholar_ids)} profil valid")

    # Langkah 3: Proses setiap profil
    final_data = []
    for sid in scholar_ids:
        profile_data = process_profile(sid)
        if profile_data and profile_data['publikasi']:
            for pub in profile_data['publikasi']:
                final_data.append({
                    'ID Scholar': sid,
                    'Nama': profile_data['nama'],
                    'Afiliasi Profil': profile_data['afiliasi'],
                    'Email': profile_data['email'],
                    'Judul Publikasi': pub['title'],
                    'Penulis': pub['authors'],
                    'Tahun': pub['year'],
                    'Journal/Conference': pub['journal'],
                    'Kesesuaian Nama': pub['nama_cocok'],
                    'Kesesuaian Afiliasi': pub['afiliasi_cocok'],
                    'Status': pub['status']
                })
        time.sleep(random.uniform(*REQUEST_DELAY))

    # Langkah 4: Simpan ke Excel
    df = pd.DataFrame(final_data)

    # Urutkan berdasarkan status meragukan
    df = df.sort_values(by=['Status', 'Kesesuaian Nama', 'Kesesuaian Afiliasi'],
                        ascending=[True, False, False])

    # Simpan ke file Excel
    output_file = "Result.xlsx"
    df.to_excel(output_file, index=False)

    # Hitung statistik
    total_pub = len(df)
    meragukan = len(df[df['Status'] == 'MERAGUKAN'])

    print("\n✅ Selesai! Hasil analisis:")
    print(f"Total Publikasi: {total_pub}")
    print(f"Publikasi Valid: {total_pub - meragukan}")
    print(f"Publikasi Meragukan: {meragukan}")
    print(f"\nFile hasil disimpan di: {output_file}")


if __name__ == "__main__":
    start()
