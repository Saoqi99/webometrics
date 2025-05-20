import time
import random
import pandas as pd
import re
from scholarly import scholarly
from rapidfuzz import fuzz, process
from scholarly import ProxyGenerator

# ========== KONFIGURASI UTAMA ========== #
TARGET_AFFILIATIONS = ["Universitas Islam Negeri Sunan Kalijaga"]
STUDENT_KEYWORDS = ['student', 'mahasiswa', 'undergraduate', 'postgraduate', 'phd candidate']
LECTURER_KEYWORDS = ['lecturer', 'dosen', 'professor', 'researcher', 'faculty', 'staff']
AFFILIATION_THRESHOLD = 75  # Tingkat kemiripan minimal
AUTHOR_MATCH_THRESHOLD = 80  # Kemiripan nama penulis
MAX_CANDIDATES = 15  # Maksimal profil kandidat
MAX_PUBLICATIONS = 100  # Maksimal publikasi per profil
REQUEST_DELAY = (1, 5)  # Jeda antar request (detik)


# ========== FUNGSI UTILITAS ========== #
def load_excluded_names(path):
    """Load daftar nama yang dikecualikan dari file Excel"""
    try:
        df = pd.read_excel(path)
        return set(df['nama'].str.lower().str.strip())
    except Exception as e:
        print(f"🚨 Gagal membaca file daftar kecualikan: {str(e)}")
        return set()


def is_student(profile):
    """Deteksi apakah profil adalah student/mahasiswa"""
    email = profile.get('email', '').lower()
    if any(s in email for s in ['student', 'mahasiswa']):
        return True

    affiliation = profile.get('affiliation', '').lower()
    return any(keyword in affiliation for keyword in STUDENT_KEYWORDS)


def is_lecturer(profile):
    """Deteksi apakah profil adalah dosen/peneliti"""
    affiliation = profile.get('affiliation', '').lower()
    return any(keyword in affiliation for keyword in LECTURER_KEYWORDS)


def build_search_queries():
    """Bangun query pencarian yang lebih spesifik"""
    base = "Universitas Islam Negeri Sunan Kalijaga"
    return [
        base,
        "UIN Sunan Kalijaga",
        f"dosen {base}",
        f"lecturer {base}",
        "Fakultas UIN Sunan Kalijaga"
    ]


# ========== FUNGSI PENCARIAN ========== #
def fetch_candidate_profiles():
    """Ambil profil kandidat dengan query spesifik"""
    all_profiles = []
    queries = build_search_queries()

    for query in queries:
        print(f"\n🔍 Mencari dengan query: '{query}'")
        try:
            search_results = scholarly.search_author(query)

            for i, author in enumerate(search_results):
                if len(all_profiles) >= MAX_CANDIDATES:
                    break

                try:
                    author = scholarly.fill(author, sections=['basics'])
                    time.sleep(random.uniform(*REQUEST_DELAY))

                    profile = {
                        'scholar_id': author['scholar_id'],
                        'name': author['name'],
                        'affiliation': author.get('affiliation', ''),
                        'email': author.get('email', ''),
                        'is_student': is_student(author),
                        'is_lecturer': is_lecturer(author)
                    }

                    # Hitung skor afiliasi
                    aff_scores = [fuzz.token_sort_ratio(profile['affiliation'].lower(), ta.lower())
                                  for ta in TARGET_AFFILIATIONS]
                    profile['aff_score'] = max(aff_scores)

                    all_profiles.append(profile)
                    print(f"  ✅ Kandidat {len(all_profiles)}: {author['name']} (Skor: {profile['aff_score']})")

                except Exception as e:
                    print(f"    🚨 Gagal memproses kandidat: {str(e)}")

        except Exception as e:
            print(f"🚨 Error saat mencari: {str(e)}")
            time.sleep(10)

    return all_profiles


def filter_profiles(candidates, excluded_names):
    """Filter kandidat berdasarkan kriteria"""
    valid_profiles = []

    for profile in candidates:
        try:
            # Skip excluded names
            if profile['name'].lower().strip() in excluded_names:
                print(f"⚠️  Dikecualikan: {profile['name']}")
                continue

            # Skip students
            if profile['is_student']:
                print(f"⚠️  Student terdeteksi: {profile['name']}")
                continue

            # Prioritize lecturers
            if not profile['is_lecturer']:
                print(f"⚠️  Bukan dosen terdeteksi: {profile['name']}")
                continue

            # Check affiliation score
            if profile['aff_score'] < AFFILIATION_THRESHOLD:
                print(f"⚠️  Afiliasi tidak cocok: {profile['name']} (Skor: {profile['aff_score']})")
                continue

            valid_profiles.append(profile)

        except Exception as e:
            print(f"⚠️  Error memproses profil: {str(e)}")

    return valid_profiles


# ========== FUNGSI PEMROSESAN ========== #
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
        author = scholarly.search_author_id(scholar_id)
        author = scholarly.fill(author, sections=['publications'])
    except Exception as e:
        print(f"   🚨 Gagal memuat profil: {str(e)}")
        return None

    results = []
    owner_name = author.get('name', '')

    for i, pub in enumerate(author['publications'][:MAX_PUBLICATIONS]):
        try:
            publication = scholarly.fill(pub)
            time.sleep(random.uniform(*REQUEST_DELAY))

            bib = publication.get('bib', {})
            authors = parse_authors(bib.get('author', ''))

            # Deteksi kecocokan
            nama_cocok = check_author_match(owner_name, authors)
            afiliasi_cocok = any(fuzz.partial_ratio(affil.lower(), bib.get('journal', '').lower()) > 75
                                 for affil in TARGET_AFFILIATIONS)

            results.append({
                'title': bib.get('title', 'No Title'),
                'authors': ', '.join(authors),
                'year': bib.get('year', ''),
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


# ========== FUNGSI UTAMA ========== #
def main():
    """Fungsi utama untuk menjalankan proses"""
    print("🚀 Memulai proses deteksi dosen dan publikasi UIN Sunan Kalijaga")

    # 1. Load daftar kecualikan
    excluded_names = load_excluded_names("daftar_kecuali.xlsx")

    # 2. Temukan dan filter profil
    print("\n🔎 Mencari kandidat dosen...")
    candidates = fetch_candidate_profiles()
    valid_profiles = filter_profiles(candidates, excluded_names)

    if not valid_profiles:
        print("\n❌ Tidak ada profil dosen yang ditemukan!")
        return

    print(f"\n🎯 Ditemukan {len(valid_profiles)} dosen valid")

    # 3. Proses setiap profil
    final_data = []
    for profile in valid_profiles:
        profile_data = process_profile(profile['scholar_id'])
        if profile_data and profile_data['publikasi']:
            for pub in profile_data['publikasi']:
                final_data.append({
                    'ID Scholar': profile_data['scholar_id'],
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

    # 4. Simpan hasil
    if final_data:
        df = pd.DataFrame(final_data)
        df = df.sort_values(by=['Status', 'Kesesuaian Nama', 'Kesesuaian Afiliasi'],
                            ascending=[True, False, False])

        output_file = "hasil_dosen_uin_suka.xlsx"
        df.to_excel(output_file, index=False)

        # Statistik
        total = len(df)
        valid = len(df[df['Status'] == 'VALID'])

        print("\n✅ Selesai! Hasil analisis:")
        print(f"Total Publikasi: {total}")
        print(f"Publikasi Valid: {valid}")
        print(f"Publikasi Meragukan: {total - valid}")
        print(f"\n💾 File hasil disimpan di: {output_file}")
    else:
        print("\n❌ Tidak ada publikasi yang ditemukan!")


if __name__ == "__main__":
    # Gunakan proxy jika diperlukan
    # pg = ProxyGenerator()
    # pg.FreeProxies()
    # scholarly.use_proxy(pg)

    main()