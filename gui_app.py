import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import random
import pandas as pd
import re
from scholarly import scholarly
from rapidfuzz import fuzz

# Konfigurasi
TARGET_AFFILIATIONS = ["Sunan Kalijaga Yogyakarta"]
AFFILIATION_THRESHOLD = 75
AUTHOR_MATCH_THRESHOLD = 80
MAX_CANDIDATES = 90
MAX_PUBLICATIONS = 75
REQUEST_DELAY = (1, 0.5)

class ScholarAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Analisis Google Scholar")
        self.setup_widgets()

    def setup_widgets(self):
        self.label = ttk.Label(self.root, text="Klik tombol untuk mulai analisis Google Scholar")
        self.label.pack(pady=10)

        self.button_start = ttk.Button(self.root, text="Mulai Analisis", command=self.run_analysis_threaded)
        self.button_start.pack(pady=5)

        self.text_output = tk.Text(self.root, height=25, width=100, wrap='word')
        self.text_output.pack(padx=10, pady=10)

        self.scrollbar = ttk.Scrollbar(self.root, command=self.text_output.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.text_output['yscrollcommand'] = self.scrollbar.set

    def log(self, message):
        self.text_output.insert(tk.END, message + "\n")
        self.text_output.see(tk.END)
        self.root.update()

    def run_analysis_threaded(self):
        threading.Thread(target=self.start_analysis).start()

    def load_excluded_names(self, path):
        try:
            df = pd.read_excel(path)
            return set(df['nama'].str.lower().str.strip())
        except Exception as e:
            self.log(f"🚨 Gagal membaca file daftar kecualikan: {str(e)}")
            return set()

    def fetch_candidate_profiles(self, query, max_results=10):
        self.log(f"\n🔍 Mencari dengan query: '{query}'")
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
                    self.log(f"  ✅ Kandidat {i + 1}: {author.get('name')}")
                    time.sleep(random.uniform(*REQUEST_DELAY))
                except Exception as e:
                    self.log(f"    🚨 Gagal memproses kandidat {i + 1}: {str(e)}")
        except Exception as e:
            self.log(f"🚨 Error saat mencari dengan query '{query}': {str(e)}")
        return candidates

    def filter_profiles(self, candidates, excluded_names):
        valid_profiles = []
        for profile in candidates:
            try:
                if profile['name'].lower().strip() in excluded_names:
                    self.log(f"⚠️  Dikecualikan: {profile['name']}")
                    continue
                aff_scores = [fuzz.token_sort_ratio(profile['affiliation'], ta.lower()) for ta in TARGET_AFFILIATIONS]
                if max(aff_scores) < AFFILIATION_THRESHOLD:
                    continue
                if 'student.uin-suka.ac.id' in profile['email']:
                    continue
                valid_profiles.append(profile)
            except:
                continue
        return valid_profiles

    def parse_authors(self, authors_str):
        if not authors_str:
            return []
        return re.split(r', | and | & ', authors_str.strip())

    def check_author_match(self, owner_name, authors_list):
        for author in authors_list:
            if fuzz.token_set_ratio(owner_name.lower(), author.lower()) >= AUTHOR_MATCH_THRESHOLD:
                return True
        return False

    def process_profile(self, scholar_id):
        self.log(f"\n📖 Memproses profil {scholar_id}")
        try:
            author = scholarly.search_author_id(scholar_id)
            author = scholarly.fill(author, sections=['publications'])
        except Exception as e:
            self.log(f"   🚨 Gagal memuat profil: {str(e)}")
            return None

        results = []
        owner_name = author.get('name', '')
        pub_count = min(len(author['publications']), MAX_PUBLICATIONS)
        self.log(f"   🔎 Memeriksa {pub_count} publikasi...")

        for i, pub in enumerate(author['publications'][:pub_count]):
            try:
                publication = scholarly.fill(pub)
                time.sleep(random.uniform(*REQUEST_DELAY))
                bib = publication.get('bib', {})
                authors = self.parse_authors(bib.get('author', ''))
                match = self.check_author_match(owner_name, authors)
                results.append({
                    'title': bib.get('title', 'No Title'),
                    'authors': ', '.join(authors),
                    'year': bib.get('year', ''),
                    'match': 'YA' if match else 'TIDAK'
                })
            except Exception as e:
                self.log(f"    🚨 Gagal memproses publikasi {i + 1}: {str(e)}")

        return {
            'scholar_id': scholar_id,
            'nama': owner_name,
            'afiliasi': author.get('affiliation', ''),
            'email': author.get('email', ''),
            'publikasi': results
        }

    def start_analysis(self):
        self.text_output.delete(1.0, tk.END)  # Bersihkan teks
        self.log("🚀 Memulai analisis...")

        excluded_names = self.load_excluded_names("daftar_kecuali.xlsx")

        all_candidates = []
        for affiliation in TARGET_AFFILIATIONS:
            candidates = self.fetch_candidate_profiles(affiliation, MAX_CANDIDATES)
            all_candidates.extend(candidates)

        valid_profiles = self.filter_profiles(all_candidates, excluded_names)
        scholar_ids = list({p['scholar_id'] for p in valid_profiles})
        self.log(f"\n🔍 Ditemukan {len(scholar_ids)} profil valid")

        final_data = []
        for sid in scholar_ids:
            profile_data = self.process_profile(sid)
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

        if final_data:
            df = pd.DataFrame(final_data)
            df.to_excel("hasil_analisis_scholar_gui.xlsx", index=False)
            self.log("\n✅ Selesai! Hasil disimpan di 'hasil_analisis_scholar_gui.xlsx'")
        else:
            self.log("⚠️ Tidak ada data yang disimpan.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScholarAnalyzerApp(root)
    root.mainloop()
