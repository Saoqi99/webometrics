import time
import random
import pandas as pd
import re
from scholarly import scholarly
from rapidfuzz import fuzz, process
from scholarly import ProxyGenerator

# pg = ProxyGenerator()
# pg.FreeProxies()  # Atau pg.SingleProxy(http="ip:port", https="ip:port")
# scholarly.use_proxy(pg)

# Konfigurasi
TARGET_AFFILIATIONS = [
    "Sunan Kalijaga Yogyakarta"
]
AFFILIATION_THRESHOLD = 75  # Tingkat kemiripan minimal untuk afiliasi
AUTHOR_MATCH_THRESHOLD = 80  # Tingkat kemiripan minimal untuk nama penulis
MAX_CANDIDATES = 200  # Maksimal profil kandidat per query
MAX_PUBLICATIONS = 75  # Maksimal publikasi yang diproses per profil
REQUEST_DELAY = (1, 5)  # Jeda acak antar request (detik)


def load_excluded_names(path):
    """Load daftar nama yang harus dikecualikan dari file Excel"""
    try:
        df = pd.read_excel(path)
        return set(df['nama'].str.lower().str.strip())
    except Exception as e:
        print(f"🚨 Gagal membaca file daftar kecualikan: {str(e)}")
        return set()


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
                # Ambil data dasar dengan retry
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


def filter_profiles(candidates, excluded_names):
    """Filter kandidat berdasarkan afiliasi, email, dan daftar kecualikan"""
    valid_profiles = []
    for profile in candidates:
        try:
            # Cek jika termasuk dalam daftar nama yang dikecualikan
            if profile['name'].lower().strip() in excluded_names:
                print(f"⚠️  Dikecualikan: {profile['name']}")
                continue

            # Cek kesamaan afiliasi
            aff_scores = [fuzz.token_sort_ratio(profile['affiliation'], ta.lower())
                          for ta in TARGET_AFFILIATIONS]
            if max(aff_scores) < AFFILIATION_THRESHOLD:
                continue

            # Cek email student
            if 'student.uin-suka.ac.id' in profile['email']:
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
    """Proses satu profil dan publikasinya"""
    print(f"\n📖 Memproses profil {scholar_id}")
    try:
        # Ambil data profil dengan retry
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
            # Ambil detail publikasi
            publication = scholarly.fill(pub)
            time.sleep(random.uniform(*REQUEST_DELAY))

            # Proses data
            bib = publication.get('bib', {})
            authors = parse_authors(bib.get('author', ''))
            match = check_author_match(owner_name, authors)

            results.append({
                'title': bib.get('title', 'No Title'),
                'authors': ', '.join(authors),
                'year': bib.get('year', ''),
                'match': 'YA' if match else 'TIDAK'
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
    # Langkah 0: Load daftar nama yang dikecualikan
    excluded_names = load_excluded_names("daftar_kecuali.xlsx")

    # Langkah 1: Ambil semua kandidat
    all_candidates = []
    for affiliation in TARGET_AFFILIATIONS:
        candidates = fetch_candidate_profiles(affiliation, MAX_CANDIDATES)
        all_candidates.extend(candidates)

    # Langkah 2: Filter kandidat
    valid_profiles = filter_profiles(all_candidates, excluded_names)
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
                    'Afiliasi': profile_data['afiliasi'],
                    'Email': profile_data['email'],
                    'Judul Publikasi': pub['title'],
                    'Penulis': pub['authors'],
                    'Tahun': pub['year'],
                    'Kecocokan': pub['match']
                })
        time.sleep(random.uniform(*REQUEST_DELAY))

    # Langkah 4: Simpan ke Excel
    df = pd.DataFrame(final_data)
    df.to_excel("hasil_analisis_scholar4.xlsx", index=False)
    print("\n✅ Selesai! Hasil disimpan di 'hasil_analisis_scholar4.xlsx'")


if __name__ == "__main__":
    start()