import time
import random
import re
import mysql.connector
import pandas as pd
from mysql.connector import Error, pooling
from scholarly import scholarly, ProxyGenerator
from rapidfuzz import fuzz, process
from tenacity import retry, stop_after_attempt, wait_exponential

# Konfigurasi Database
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'webometrics',
    'port': 3306,
    'pool_name': 'webometrics_pool',
    'pool_size': 5
}

# Konfigurasi Aplikasi
TARGET_AFFILIATIONS = ["Sunan Kalijaga Yogyakarta"]
AFFILIATION_THRESHOLD = 75
AUTHOR_MATCH_THRESHOLD = 80
MAX_CANDIDATES = 30
MAX_PUBLICATIONS = 75
REQUEST_DELAY = (1, 3)


class DatabaseHandler:
    def __init__(self):
        self.connection_pool = self.create_connection_pool()
        self.initialize_database()

    def create_connection_pool(self):
        try:
            return pooling.MySQLConnectionPool(**DB_CONFIG)
        except Error as e:
            print(f"Error creating connection pool: {e}")
            raise

    def get_connection(self):
        return self.connection_pool.get_connection()

    def initialize_database(self):
        init_queries = [
            """CREATE TABLE IF NOT EXISTS profiles (
                id INT AUTO_INCREMENT PRIMARY KEY,
                scholar_id VARCHAR(255) UNIQUE,
                nama VARCHAR(255) NOT NULL,
                afiliasi VARCHAR(255),
                email VARCHAR(255),
                terakhir_diperbarui TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",

            """CREATE TABLE IF NOT EXISTS publikasi (
                id INT AUTO_INCREMENT PRIMARY KEY,
                profile_id INT NOT NULL,
                judul TEXT NOT NULL,
                penulis TEXT,
                tahun YEAR,
                kecocokan ENUM('YA', 'TIDAK'),
                sumber_url VARCHAR(255) UNIQUE,
                FOREIGN KEY (profile_id) REFERENCES profiles(id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",

            "CREATE INDEX IF NOT EXISTS idx_nama ON profiles(nama)",
            "CREATE INDEX IF NOT EXISTS idx_tahun ON publikasi(tahun)"
        ]

        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            for query in init_queries:
                cursor.execute(query)
            conn.commit()
        except Error as e:
            print(f"Error initializing database: {e}")
            conn.rollback()
        finally:
            cursor.close()
            conn.close()
    def simpan_publikasi(self, profile_id, publikasi):
        query = """
        INSERT IGNORE INTO publikasi 
            (profile_id, judul, penulis, tahun, kecocokan, sumber_url)
        VALUES (%s, %s, %s, %s, %s, %s)
        """
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            for pub in publikasi:
                # Cek duplikasi fuzzy match
                if not self.cek_duplikasi_publikasi(profile_id, pub['Judul Publikasi']):
                    cursor.execute(query, (
                        profile_id,
                        pub['Judul Publikasi'],
                        pub['Penulis'],
                        pub['Tahun'],
                        pub['Kecocokan'],
                        pub['sumber_url']
                    ))
            conn.commit()
        except Error as e:
            print(f"Error saving publications: {e}")
        finally:
            cursor.close()
            conn.close()

    def cek_duplikasi_publikasi(self, profile_id, judul):
        query = """
        SELECT judul FROM publikasi
        WHERE profile_id = %s
        AND SOUNDEX(judul) = SOUNDEX(%s)
        """
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute(query, (profile_id, judul))
            return cursor.fetchone() is not None
        except Error as e:
            print(f"Error checking duplicate: {e}")
            return False
        finally:
            cursor.close()
            conn.close()


class ScholarAnalyzer:
    def __init__(self):
        self.db_handler = DatabaseHandler()
        self.processed_profiles = set()
        self._load_existing_profiles()
        self._configure_scholarly()

    def _configure_scholarly(self):
        # Konfigurasi user agent dan headers
        scholarly.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://scholar.google.com/'
        }
        scholarly.timeout = 30

    def _log(self, message):
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}")

    def _load_existing_profiles(self):
        conn = self.db_handler.get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT scholar_id FROM profiles")
            self.processed_profiles = {row[0] for row in cursor.fetchall()}
        except Error as e:
            self._log(f"Error loading existing profiles: {e}")
        finally:
            cursor.close()
            conn.close()

    def _delay(self):
        time.sleep(random.uniform(*REQUEST_DELAY))

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
    def fetch_candidate_profiles(self, query):
        self._log(f"Mencari kandidat untuk: {query}")
        try:
            search_results = scholarly.search_author(query)
            return [
                self._process_candidate(author)
                for i, author in enumerate(search_results)
                if i < MAX_CANDIDATES
            ]
        except Exception as e:
            self._log(f"Error fetching candidates: {e}")
            return []

    def _process_candidate(self, author):
        self._delay()
        try:
            return {
                'scholar_id': author.get('scholar_id'),
                'nama': author.get('name'),
                'afiliasi': author.get('affiliation', '').lower(),
                'email': author.get('email', '').lower()
            }
        except Exception as e:
            self._log(f"Error processing candidate: {e}")
            return None

    def process_profile(self, scholar_id):
        self._log(f"Memproses profil: {scholar_id}")
        try:
            # Gunakan metode yang benar untuk mencari author
            author = scholarly.search_by_author_id(scholar_id)
            author = scholarly.fill(author, sections=['publications'])
            return self._process_publications(author)
        except Exception as e:
            self._log(f"Error processing profile: {e}")
            return None

    def _process_publications(self, author):
        profile_data = {
            'scholar_id': author.get('scholar_id'),
            'nama': author.get('name', ''),
            'afiliasi': author.get('affiliation', ''),
            'email': author.get('email', '')
        }

        profile_id = self.db_handler.simpan_profil(profile_data)
        if not profile_id:
            return None

        publications = []
        for pub in author['publications'][:MAX_PUBLICATIONS]:
            try:
                publication = scholarly.fill(pub)
                publications.append(self._parse_publication(publication, profile_data['nama']))
                self._delay()
            except Exception as e:
                self._log(f"Error processing publication: {e}")

        if publications:
            self.db_handler.simpan_publikasi(profile_id, publications)

        return publications

    def _parse_publication(self, publication, owner_name):
        bib = publication.get('bib', {})
        authors = self._parse_authors(bib.get('author', ''))
        return {
            'Judul Publikasi': bib.get('title', 'No Title'),
            'Penulis': ', '.join(authors),
            'Tahun': bib.get('year', ''),
            'Kecocokan': 'YA' if self._check_author_match(owner_name, authors) else 'TIDAK',
            'sumber_url': publication.get('url_scholarbib', '')
        }

    def _parse_authors(self, authors_str):
        return re.split(r', | and | & ', authors_str.strip()) if authors_str else []

    def _check_author_match(self, owner_name, authors):
        return any(fuzz.token_set_ratio(owner_name.lower(), a.lower()) >= AUTHOR_MATCH_THRESHOLD for a in authors)

    def generate_report(self):
        conn = self.db_handler.get_connection()
        try:
            query = """
            SELECT p.nama, p.afiliasi, pu.judul, pu.penulis, pu.tahun, pu.kecocokan
            FROM publikasi pu
            JOIN profiles p ON pu.profile_id = p.id
            """
            df = pd.read_sql(query, conn)
            df.to_excel("laporan_analisis.xlsx", index=False)
            self._log("Laporan berhasil di-generate")
        except Error as e:
            self._log(f"Error generating report: {e}")
        finally:
            conn.close()


def main():
    analyzer = ScholarAnalyzer()

    try:
        # Langkah 1: Kumpulkan kandidat
        all_candidates = []
        for affiliation in TARGET_AFFILIATIONS:
            candidates = analyzer.fetch_candidate_profiles(affiliation)
            all_candidates.extend(candidates)

        # Langkah 2: Filter kandidat
        valid_profiles = analyzer.filter_profiles(all_candidates)

        # Langkah 3: Proses profil
        for profile in valid_profiles:
            if profile['scholar_id'] not in analyzer.processed_profiles:
                analyzer.process_profile(profile['scholar_id'])

        # Langkah 4: Generate laporan
        analyzer.generate_report()

    except KeyboardInterrupt:
        analyzer._log("Proses dihentikan oleh pengguna")
    finally:
        analyzer._log("Analisis selesai")


if __name__ == "__main__":
    analyzer = ScholarAnalyzer()

    try:
        # Langkah 1: Kumpulkan kandidat
        all_candidates = []
        for affiliation in TARGET_AFFILIATIONS:
            candidates = analyzer.fetch_candidate_profiles(affiliation)
            all_candidates.extend(candidates)

        # Langkah 2: Filter kandidat
        valid_profiles = analyzer.filter_profiles(all_candidates)

        # Langkah 3: Proses profil
        for profile in valid_profiles:
            if profile['scholar_id'] not in analyzer.processed_profiles:
                analyzer.process_profile(profile['scholar_id'])

        # Langkah 4: Generate laporan
        analyzer.generate_report()

    except KeyboardInterrupt:
        analyzer._log("Proses dihentikan oleh pengguna")
    finally:
        analyzer._log("Analisis selesai")