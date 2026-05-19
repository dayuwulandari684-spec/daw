from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
from docx.oxml import OxmlElement

NAMA_PENULIS = "ROBY PRASETYO ADI"
MAPEL = "Pendidikan Kewarganegaraan (PKN)"
KELAS = "XI"
REF = "Buku Siswa PKN Kelas XI, Kemendikbud RI; Modul PKN Merdeka Belajar Kelas XI; UUD NRI 1945"

pg_data = [
    # ---- MATERI 1: ATHG TERHADAP IDEOLOGI PANCASILA (Soal 1-7) ----
    {
        "no": 1, "elemen": "Ancaman, Tantangan, Hambatan, dan Gangguan (ATHG) terhadap Ideologi Pancasila",
        "materi": "Pengertian ATHG terhadap Ideologi Pancasila",
        "cp": "Peserta didik mampu menganalisis berbagai bentuk ancaman, tantangan, hambatan, dan gangguan terhadap ideologi Pancasila",
        "ipk": "Mengidentifikasi perbedaan antara ancaman, tantangan, hambatan, dan gangguan terhadap Pancasila",
        "indikator": "Disajikan deskripsi suatu kondisi yang mempengaruhi ideologi Pancasila, peserta didik dapat mengklasifikasikan jenis ATHG yang dimaksud",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Sebuah gerakan yang secara terang-terangan berupaya menggantikan Pancasila dengan ideologi lain melalui tindakan kekerasan dan pemberontakan bersenjata merupakan salah satu bentuk ancaman nyata bagi ideologi Pancasila. Berdasarkan pernyataan tersebut, yang dimaksud dengan ANCAMAN terhadap ideologi Pancasila adalah ...",
        "A": "Upaya yang berasal dari dalam negeri untuk menyempurnakan Pancasila agar lebih relevan",
        "B": "Kondisi yang secara tidak sengaja menghambat pelaksanaan nilai-nilai Pancasila",
        "C": "Setiap usaha dan kegiatan baik dari dalam maupun luar negeri yang dinilai membahayakan kelangsungan ideologi Pancasila",
        "D": "Permasalahan sosial yang muncul akibat perbedaan penafsiran terhadap nilai Pancasila",
        "E": "Ketidakmampuan aparatur negara dalam menjalankan amanat Pancasila secara konsisten",
        "jawaban": "C",
    },
    {
        "no": 2, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Bentuk Ancaman terhadap Pancasila dari Dalam Negeri",
        "cp": "Peserta didik mampu menganalisis berbagai bentuk ancaman terhadap ideologi Pancasila dari dalam negeri",
        "ipk": "Mengidentifikasi contoh ancaman ideologi terhadap Pancasila yang berasal dari dalam negeri",
        "indikator": "Diberikan beberapa pernyataan tentang ancaman, peserta didik dapat menentukan ancaman ideologi terhadap Pancasila yang berasal dari dalam negeri",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Perhatikan pernyataan-pernyataan berikut:\n1. Gerakan separatisme yang ingin memisahkan diri dari NKRI\n2. Penyebaran paham radikalisme dan ekstremisme\n3. Masuknya budaya asing yang bertentangan dengan nilai Pancasila\n4. Munculnya gerakan komunisme yang ingin mengubah dasar negara\n5. Tekanan dari lembaga keuangan internasional yang mengancam kedaulatan ekonomi\nYang termasuk ancaman terhadap ideologi Pancasila yang berasal dari DALAM NEGERI adalah ...",
        "A": "1, 2, dan 3",
        "B": "1, 2, dan 4",
        "C": "2, 3, dan 5",
        "D": "3, 4, dan 5",
        "E": "1, 4, dan 5",
        "jawaban": "B",
    },
    {
        "no": 3, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Ancaman terhadap Pancasila dari Luar Negeri",
        "cp": "Peserta didik mampu menganalisis ancaman terhadap ideologi Pancasila yang bersumber dari luar negeri",
        "ipk": "Mengidentifikasi bentuk ancaman ideologi terhadap Pancasila yang bersumber dari luar negeri",
        "indikator": "Disajikan contoh situasi global, peserta didik dapat mengidentifikasi ancaman terhadap Pancasila yang bersumber dari luar negeri",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Penyebaran ideologi liberal-kapitalis melalui globalisasi yang mendorong gaya hidup individualistis dan mengikis nilai-nilai gotong royong dan kekeluargaan yang merupakan bagian dari nilai Pancasila merupakan contoh ancaman Pancasila yang bersumber dari ...",
        "A": "Dalam negeri berupa pergeseran budaya lokal",
        "B": "Luar negeri berupa penetrasi ideologi asing",
        "C": "Dalam negeri berupa kemunduran moral bangsa",
        "D": "Konflik antar kelompok masyarakat di Indonesia",
        "E": "Kelemahan sistem pendidikan nasional Indonesia",
        "jawaban": "B",
    },
    {
        "no": 4, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Tantangan terhadap Pancasila di Era Globalisasi",
        "cp": "Peserta didik mampu menganalisis tantangan terhadap eksistensi Pancasila di era globalisasi",
        "ipk": "Menjelaskan tantangan yang dihadapi Pancasila sebagai ideologi bangsa di era globalisasi",
        "indikator": "Diberikan gambaran kondisi era globalisasi, peserta didik dapat menganalisis tantangan yang dihadapi Pancasila sebagai ideologi bangsa",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Di era globalisasi, informasi dan budaya dari seluruh penjuru dunia dapat dengan mudah masuk ke Indonesia melalui internet dan media sosial. Kondisi ini menjadi tantangan bagi ideologi Pancasila karena ...",
        "A": "Globalisasi membuat Indonesia semakin diakui oleh dunia internasional",
        "B": "Masyarakat Indonesia menjadi lebih modern dan maju dalam segala bidang",
        "C": "Nilai-nilai asing yang bertentangan dengan Pancasila lebih mudah mempengaruhi cara berpikir masyarakat Indonesia",
        "D": "Pemerintah Indonesia kesulitan membangun infrastruktur teknologi yang memadai",
        "E": "Indonesia harus menyesuaikan diri dengan standar internasional dalam semua aspek kehidupan",
        "jawaban": "C",
    },
    {
        "no": 5, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Hambatan terhadap Pelaksanaan Nilai-nilai Pancasila",
        "cp": "Peserta didik mampu mengidentifikasi hambatan dalam pelaksanaan nilai-nilai Pancasila dalam kehidupan berbangsa",
        "ipk": "Membedakan hambatan internal dan eksternal dalam pelaksanaan nilai-nilai Pancasila",
        "indikator": "Disajikan kondisi dalam kehidupan berbangsa, peserta didik dapat mengidentifikasi yang merupakan hambatan dalam pelaksanaan nilai Pancasila",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Rendahnya pemahaman masyarakat terhadap nilai-nilai Pancasila akibat minimnya pendidikan karakter di sekolah dan kurangnya keteladanan para pemimpin dalam mengamalkan Pancasila merupakan salah satu bentuk ...",
        "A": "Ancaman terhadap Pancasila dari luar negeri",
        "B": "Gangguan terhadap stabilitas nasional",
        "C": "Hambatan dalam pelaksanaan nilai-nilai Pancasila",
        "D": "Tantangan eksternal bagi eksistensi Pancasila",
        "E": "Ancaman militer terhadap kedaulatan negara",
        "jawaban": "C",
    },
    {
        "no": 6, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Upaya Menghadapi ATHG terhadap Pancasila",
        "cp": "Peserta didik mampu merumuskan upaya-upaya untuk menghadapi ATHG terhadap ideologi Pancasila",
        "ipk": "Menganalisis upaya yang tepat untuk menghadapi ancaman terhadap ideologi Pancasila",
        "indikator": "Diberikan berbagai upaya menghadapi ATHG, peserta didik dapat menentukan upaya yang paling tepat dan efektif",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "Untuk menghadapi penyebaran paham radikalisme dan separatisme yang mengancam ideologi Pancasila, upaya yang paling tepat dilakukan oleh warga negara dalam kehidupan sehari-hari adalah ...",
        "A": "Menutup semua akses internet agar paham radikalisme tidak bisa masuk",
        "B": "Membiarkan pemerintah yang menangani karena itu bukan urusan warga biasa",
        "C": "Memperkuat pemahaman nilai-nilai Pancasila, aktif melaporkan konten radikal, dan membangun toleransi antar sesama warga",
        "D": "Menjauhi semua orang yang berbeda agama, suku, dan pandangan politik",
        "E": "Menolak segala bentuk pengaruh asing agar budaya Indonesia tetap murni",
        "jawaban": "C",
    },
    {
        "no": 7, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Peran Bela Negara dalam Menghadapi ATHG",
        "cp": "Peserta didik mampu menganalisis peran bela negara dalam menghadapi ATHG terhadap Pancasila",
        "ipk": "Menghubungkan konsep bela negara dengan upaya menghadapi ATHG terhadap Pancasila",
        "indikator": "Diberikan konsep bela negara, peserta didik dapat menganalisis hubungannya dengan upaya menghadapi ATHG terhadap Pancasila",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "Bela negara bukan hanya dilakukan melalui perjuangan fisik di medan perang, tetapi juga melalui hal-hal yang dilakukan dalam kehidupan sehari-hari. Contoh bela negara yang dapat dilakukan seorang pelajar dalam menghadapi ancaman terhadap ideologi Pancasila adalah ...",
        "A": "Mengikuti latihan militer dan mempelajari cara menggunakan senjata",
        "B": "Belajar dengan sungguh-sungguh, mengamalkan nilai Pancasila, dan menyebarkan konten positif di media sosial",
        "C": "Menghindari semua kegiatan yang berhubungan dengan politik dan pemerintahan",
        "D": "Menggalang dana untuk mendukung operasi militer di perbatasan negara",
        "E": "Memilih untuk tinggal di daerah terpencil agar terhindar dari pengaruh asing",
        "jawaban": "B",
    },
    # ---- MATERI 2: BENTUK NEGARA, PEMERINTAHAN, SISTEM PEMERINTAHAN (Soal 8-14) ----
    {
        "no": 8, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Bentuk Negara: Negara Kesatuan dan Negara Federal",
        "cp": "Peserta didik mampu membedakan berbagai bentuk negara dan bentuk pemerintahan",
        "ipk": "Membedakan ciri-ciri negara kesatuan dan negara federal",
        "indikator": "Disajikan ciri-ciri bentuk negara, peserta didik dapat membedakan negara kesatuan dengan negara federal",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Negara Kesatuan Republik Indonesia (NKRI) menerapkan bentuk negara kesatuan. Berikut ini yang merupakan ciri utama dari NEGARA KESATUAN adalah ...",
        "A": "Terdapat dua tingkat pemerintahan yang masing-masing memiliki kedaulatan sendiri",
        "B": "Setiap daerah memiliki konstitusi sendiri yang berbeda dengan konstitusi pusat",
        "C": "Hanya ada satu pemerintah pusat yang memegang kedaulatan penuh atas seluruh wilayah negara",
        "D": "Kepala negara dipilih melalui pemilihan oleh parlemen dari setiap negara bagian",
        "E": "Daerah-daerah berhak untuk memisahkan diri jika tidak puas dengan kebijakan pusat",
        "jawaban": "C",
    },
    {
        "no": 9, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Bentuk Pemerintahan: Monarki dan Republik",
        "cp": "Peserta didik mampu membedakan bentuk pemerintahan monarki dan republik beserta contoh negaranya",
        "ipk": "Membedakan ciri-ciri pemerintahan monarki dan republik",
        "indikator": "Diberikan deskripsi bentuk pemerintahan suatu negara, peserta didik dapat mengidentifikasi apakah termasuk monarki atau republik",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Suatu negara dipimpin oleh seorang kepala negara yang mendapat jabatannya melalui pewarisan turun-temurun (keturunan). Kepala negara tersebut memegang jabatan seumur hidup. Bentuk pemerintahan negara tersebut adalah ...",
        "A": "Republik",
        "B": "Demokrasi",
        "C": "Oligarki",
        "D": "Monarki",
        "E": "Teokrasi",
        "jawaban": "D",
    },
    {
        "no": 10, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Sistem Pemerintahan Presidensial",
        "cp": "Peserta didik mampu menganalisis ciri-ciri sistem pemerintahan presidensial",
        "ipk": "Mengidentifikasi ciri-ciri utama sistem pemerintahan presidensial",
        "indikator": "Diberikan pernyataan tentang sistem pemerintahan, peserta didik dapat mengidentifikasi ciri-ciri sistem pemerintahan presidensial",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Indonesia menganut sistem pemerintahan presidensial. Berikut ini yang BUKAN merupakan ciri-ciri sistem pemerintahan presidensial adalah ...",
        "A": "Presiden berkedudukan sebagai kepala negara sekaligus kepala pemerintahan",
        "B": "Kabinet (menteri) bertanggung jawab kepada presiden, bukan kepada parlemen",
        "C": "Presiden dipilih langsung oleh rakyat melalui pemilihan umum",
        "D": "Pemerintah (eksekutif) dapat dijatuhkan oleh parlemen melalui mosi tidak percaya",
        "E": "Masa jabatan presiden ditentukan secara pasti dalam konstitusi",
        "jawaban": "D",
    },
    {
        "no": 11, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Sistem Pemerintahan Parlementer",
        "cp": "Peserta didik mampu membedakan sistem pemerintahan presidensial dan parlementer",
        "ipk": "Membedakan ciri-ciri sistem pemerintahan parlementer dengan presidensial",
        "indikator": "Diberikan pernyataan tentang sistem pemerintahan, peserta didik dapat membedakan sistem parlementer dengan presidensial",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Perhatikan ciri-ciri sistem pemerintahan berikut:\n1. Kepala pemerintahan adalah Perdana Menteri\n2. Kabinet bertanggung jawab kepada parlemen\n3. Pemerintah dapat dibubarkan oleh parlemen melalui mosi tidak percaya\n4. Kepala negara dan kepala pemerintahan dijabat oleh orang yang sama\n5. Parlemen dapat dibubarkan oleh kepala negara\nCiri-ciri yang merupakan sistem pemerintahan PARLEMENTER adalah ...",
        "A": "1, 2, dan 3",
        "B": "1, 3, dan 4",
        "C": "2, 4, dan 5",
        "D": "1, 2, dan 5",
        "E": "3, 4, dan 5",
        "jawaban": "D",
    },
    {
        "no": 12, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Bentuk Negara dan Sistem Pemerintahan Indonesia",
        "cp": "Peserta didik mampu menjelaskan bentuk negara dan sistem pemerintahan Indonesia berdasarkan UUD 1945",
        "ipk": "Menjelaskan bentuk negara, bentuk pemerintahan, dan sistem pemerintahan Indonesia berdasarkan UUD NRI 1945",
        "indikator": "Diberikan pertanyaan tentang Indonesia, peserta didik dapat menjelaskan bentuk negara dan sistem pemerintahan Indonesia berdasarkan UUD 1945",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Berdasarkan UUD NRI 1945, Indonesia memiliki bentuk negara, bentuk pemerintahan, dan sistem pemerintahan yang jelas. Pasangan yang tepat untuk melengkapi pernyataan tersebut adalah ...",
        "A": "Negara federal – Monarki konstitusional – Parlementer",
        "B": "Negara kesatuan – Republik – Presidensial",
        "C": "Negara federal – Republik – Presidensial",
        "D": "Negara kesatuan – Monarki – Parlementer",
        "E": "Negara serikat – Republik – Semi-presidensial",
        "jawaban": "B",
    },
    {
        "no": 13, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Otonomi Daerah dalam Sistem Pemerintahan Indonesia",
        "cp": "Peserta didik mampu menganalisis penerapan otonomi daerah dalam sistem pemerintahan NKRI",
        "ipk": "Menjelaskan konsep otonomi daerah dan hubungannya dengan bentuk negara kesatuan Indonesia",
        "indikator": "Diberikan pernyataan tentang otonomi daerah, peserta didik dapat menjelaskan hubungannya dengan sistem pemerintahan NKRI",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Indonesia sebagai negara kesatuan memberikan otonomi kepada daerah-daerah untuk mengatur rumah tangganya sendiri. Pemberian otonomi daerah ini bertujuan untuk ...",
        "A": "Memberikan kebebasan daerah untuk membuat konstitusi sendiri yang berbeda dengan UUD 1945",
        "B": "Mempersiapkan daerah untuk menjadi negara bagian yang berdiri sendiri",
        "C": "Mempercepat pembangunan, meningkatkan pelayanan publik, dan memberdayakan masyarakat daerah",
        "D": "Mengurangi peran pemerintah pusat dalam semua urusan pemerintahan",
        "E": "Memberikan hak kepada daerah untuk menentukan kepala negara sendiri",
        "jawaban": "C",
    },
    {
        "no": 14, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Kedaulatan Rakyat dalam Sistem Pemerintahan Indonesia",
        "cp": "Peserta didik mampu menjelaskan konsep kedaulatan rakyat dalam sistem pemerintahan Indonesia",
        "ipk": "Menghubungkan konsep kedaulatan rakyat dengan pelaksanaan sistem pemerintahan di Indonesia",
        "indikator": "Diberikan pernyataan tentang kedaulatan, peserta didik dapat menghubungkan konsep kedaulatan rakyat dengan sistem pemerintahan Indonesia",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Pasal 1 ayat (2) UUD NRI 1945 menyatakan bahwa 'Kedaulatan berada di tangan rakyat dan dilaksanakan menurut Undang-Undang Dasar.' Pernyataan ini berarti ...",
        "A": "Rakyat secara langsung menjalankan semua urusan pemerintahan setiap saat",
        "B": "Kedaulatan rakyat dilaksanakan melalui lembaga-lembaga negara yang diatur dalam UUD 1945",
        "C": "Rakyat tidak perlu ikut campur dalam urusan pemerintahan karena sudah ada lembaga negara",
        "D": "Pemerintah adalah satu-satunya pemegang kekuasaan tertinggi di Indonesia",
        "E": "Kedaulatan hanya berlaku saat pelaksanaan pemilihan umum saja",
        "jawaban": "B",
    },
    # ---- MATERI 3: SIKAP WN TERHADAP SISTEM PEMERINTAHAN RI (Soal 15-20) ----
    {
        "no": 15, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Partisipasi Warga Negara dalam Sistem Pemerintahan",
        "cp": "Peserta didik mampu menganalisis bentuk-bentuk partisipasi warga negara dalam pelaksanaan sistem pemerintahan RI",
        "ipk": "Mengidentifikasi bentuk partisipasi aktif warga negara dalam sistem pemerintahan RI",
        "indikator": "Diberikan berbagai contoh perilaku warga negara, peserta didik dapat mengidentifikasi bentuk partisipasi aktif dalam sistem pemerintahan",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Sebagai warga negara yang baik, seseorang dituntut untuk berpartisipasi aktif dalam sistem pemerintahan Republik Indonesia. Berikut ini yang merupakan contoh partisipasi aktif warga negara dalam sistem pemerintahan adalah ...",
        "A": "Menghindari semua urusan politik karena dianggap kotor dan merugikan",
        "B": "Menggunakan hak pilih dalam Pemilu, membayar pajak, dan aktif mengawasi jalannya pemerintahan",
        "C": "Hanya mematuhi peraturan yang dianggap menguntungkan diri sendiri",
        "D": "Menyerahkan sepenuhnya semua urusan negara kepada para wakil rakyat",
        "E": "Melakukan aksi unjuk rasa setiap saat sebagai bentuk kontrol terhadap pemerintah",
        "jawaban": "B",
    },
    {
        "no": 16, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Hak dan Kewajiban Warga Negara dalam Sistem Pemerintahan",
        "cp": "Peserta didik mampu membedakan hak dan kewajiban warga negara dalam sistem pemerintahan RI",
        "ipk": "Membedakan hak dan kewajiban warga negara dalam pelaksanaan sistem pemerintahan RI",
        "indikator": "Disajikan pernyataan tentang hak dan kewajiban, peserta didik dapat membedakan yang termasuk hak dan yang termasuk kewajiban warga negara",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Dalam sistem pemerintahan Republik Indonesia, warga negara memiliki hak dan kewajiban yang diatur dalam UUD NRI 1945. Berikut ini yang merupakan KEWAJIBAN warga negara berdasarkan UUD NRI 1945 adalah ...",
        "A": "Mendapatkan perlindungan hukum yang sama",
        "B": "Memperoleh pendidikan yang layak dari negara",
        "C": "Mendapatkan pekerjaan dan penghidupan yang layak",
        "D": "Membela negara dan membayar pajak sesuai undang-undang",
        "E": "Bebas mengeluarkan pendapat dan berekspresi",
        "jawaban": "D",
    },
    {
        "no": 17, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Sikap Kritis dan Konstruktif terhadap Pemerintahan",
        "cp": "Peserta didik mampu menunjukkan sikap kritis dan konstruktif terhadap pelaksanaan sistem pemerintahan RI",
        "ipk": "Menerapkan sikap kritis dan konstruktif yang tepat terhadap kebijakan pemerintah",
        "indikator": "Diberikan situasi ketidakpuasan terhadap kebijakan pemerintah, peserta didik dapat menentukan sikap kritis yang konstruktif dan sesuai dengan norma",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "Seorang warga negara tidak setuju dengan kebijakan pemerintah yang dinilai merugikan masyarakat. Sikap yang paling tepat dan konstruktif yang dapat ditunjukkan warga negara tersebut adalah ...",
        "A": "Melakukan aksi anarki dan perusakan fasilitas umum sebagai bentuk protes",
        "B": "Menyebarkan informasi palsu (hoaks) tentang pemerintah di media sosial untuk menarik simpati",
        "C": "Diam dan menerima saja karena pemerintah pasti lebih tahu yang terbaik",
        "D": "Menyampaikan kritik dan aspirasi melalui jalur yang sah seperti forum musrenbang, kepada wakil rakyat, atau media yang bertanggung jawab",
        "E": "Mengajak masyarakat untuk tidak membayar pajak sebagai bentuk protes terhadap pemerintah",
        "jawaban": "D",
    },
    {
        "no": 18, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Peran Serta Warga Negara dalam Pengawasan Pemerintahan",
        "cp": "Peserta didik mampu menganalisis peran serta warga negara dalam mengawasi jalannya pemerintahan",
        "ipk": "Menjelaskan mekanisme pengawasan warga negara terhadap penyelenggaraan pemerintahan",
        "indikator": "Diberikan kasus penyelenggaraan pemerintahan, peserta didik dapat menjelaskan bentuk pengawasan yang dapat dilakukan warga negara",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Pengawasan terhadap jalannya pemerintahan merupakan hal yang penting dalam sistem demokrasi. Berikut ini yang BUKAN merupakan mekanisme pengawasan yang dapat dilakukan oleh warga negara terhadap penyelenggaraan pemerintahan adalah ...",
        "A": "Menyampaikan laporan kepada Komisi Pemberantasan Korupsi (KPK) jika mengetahui dugaan korupsi",
        "B": "Mengikuti pemilihan umum dan memilih wakil rakyat yang jujur dan kompeten",
        "C": "Mengikuti partai politik dan berpartisipasi dalam proses legislasi",
        "D": "Melakukan main hakim sendiri kepada pejabat yang diduga melakukan korupsi",
        "E": "Aktif dalam organisasi masyarakat sipil yang memantau kinerja pemerintah",
        "jawaban": "D",
    },
    {
        "no": 19, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Sikap Taat Hukum sebagai Warga Negara",
        "cp": "Peserta didik mampu menunjukkan sikap taat hukum sebagai warga negara dalam kehidupan bermasyarakat",
        "ipk": "Menerapkan sikap taat hukum sebagai perwujudan dukungan terhadap sistem pemerintahan RI",
        "indikator": "Diberikan situasi kehidupan sehari-hari, peserta didik dapat menentukan sikap taat hukum yang mencerminkan dukungan terhadap sistem pemerintahan RI",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Sikap taat hukum merupakan salah satu wujud nyata dukungan warga negara terhadap sistem pemerintahan RI. Berikut ini yang merupakan contoh sikap taat hukum dalam kehidupan sehari-hari adalah ...",
        "A": "Mematuhi peraturan lalu lintas hanya jika ada polisi yang mengawasi",
        "B": "Membayar pajak, mematuhi rambu-rambu lalu lintas, dan mengikuti aturan yang berlaku di masyarakat",
        "C": "Memanfaatkan celah hukum untuk kepentingan pribadi selama tidak tertangkap",
        "D": "Mengikuti hukum hanya untuk hal-hal yang dianggap penting saja",
        "E": "Memilih untuk tinggal di luar negeri agar terhindar dari aturan yang memberatkan",
        "jawaban": "B",
    },
    {
        "no": 20, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Peran Pemuda dalam Sistem Pemerintahan RI",
        "cp": "Peserta didik mampu menganalisis peran pemuda/pelajar dalam mendukung pelaksanaan sistem pemerintahan RI",
        "ipk": "Menunjukkan peran nyata pemuda/pelajar dalam mendukung pelaksanaan sistem pemerintahan RI",
        "indikator": "Diberikan contoh perilaku pemuda, peserta didik dapat menganalisis peran nyata pemuda dalam mendukung pelaksanaan sistem pemerintahan RI",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "Sebagai generasi penerus bangsa, pelajar memiliki peran penting dalam mendukung pelaksanaan sistem pemerintahan Republik Indonesia. Bentuk peran nyata pelajar dalam mendukung sistem pemerintahan yang baik (good governance) adalah ...",
        "A": "Menunggu sampai dewasa dan memiliki jabatan sebelum berkontribusi pada negara",
        "B": "Menolak mempelajari nilai-nilai Pancasila karena dianggap tidak relevan dengan kehidupan modern",
        "C": "Berprestasi di bidang akademik dan non-akademik, anti korupsi, jujur, dan ikut menjaga kondusivitas lingkungan",
        "D": "Hanya fokus pada kepentingan akademik dan tidak perlu peduli dengan urusan pemerintahan",
        "E": "Aktif mendukung satu partai politik tertentu dan menjadi juru kampanye sejak usia dini",
        "jawaban": "C",
    },
]

essay_data = [
    {
        "no": 1, "elemen": "Ancaman, Tantangan, Hambatan, dan Gangguan (ATHG) terhadap Ideologi Pancasila",
        "materi": "Pengertian dan Klasifikasi ATHG terhadap Pancasila",
        "cp": "Peserta didik mampu menganalisis berbagai bentuk ATHG terhadap ideologi Pancasila secara komprehensif",
        "ipk": "Menjelaskan pengertian dan perbedaan antara ancaman, tantangan, hambatan, dan gangguan terhadap Pancasila",
        "indikator": "Disajikan konsep ATHG, peserta didik dapat menjelaskan pengertian dan perbedaan masing-masing ATHG beserta contohnya",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Pancasila sebagai dasar negara dan ideologi bangsa Indonesia menghadapi berbagai persoalan yang dapat mengancam eksistensinya. Dalam konteks ketahanan nasional, para ahli membedakan persoalan-persoalan tersebut menjadi empat kategori, yaitu Ancaman, Tantangan, Hambatan, dan Gangguan (ATHG).\n\nSetiap kategori memiliki karakteristik dan tingkat bahaya yang berbeda-beda terhadap Pancasila. Memahami perbedaan di antara keempatnya sangat penting agar masyarakat dapat menentukan sikap dan tindakan yang tepat dalam menghadapinya.\n\nSebagai pelajar yang hidup di era modern dengan berbagai dinamika sosial, politik, dan teknologi yang kompleks, pemahaman mendalam tentang ATHG terhadap Pancasila menjadi bekal penting untuk ikut menjaga keutuhan ideologi bangsa.",
        "pertanyaan": "Jelaskan pengertian dan perbedaan antara Ancaman, Tantangan, Hambatan, dan Gangguan (ATHG) terhadap ideologi Pancasila! Berikan masing-masing satu contoh konkret untuk setiap kategori!",
        "jawaban": "Pengertian dan Perbedaan ATHG terhadap Ideologi Pancasila:\n\n1. ANCAMAN\nPengertian: Setiap usaha dan kegiatan, baik dari dalam maupun luar negeri, yang dinilai membahayakan kedaulatan negara, keutuhan wilayah, dan keselamatan bangsa termasuk eksistensi Pancasila.\nContoh: Gerakan separatisme bersenjata yang ingin menggantikan Pancasila dengan ideologi lain, atau pemberontakan PKI yang ingin mengubah dasar negara.\n\n2. TANTANGAN\nPengertian: Hal atau upaya yang bertujuan menggugah kemampuan bangsa untuk menghadapi persoalan-persoalan yang ada, bersifat dinamis dan merupakan konsekuensi dari perkembangan zaman.\nContoh: Derasnya arus globalisasi dan kemajuan teknologi informasi yang membawa nilai-nilai asing yang bertentangan dengan Pancasila.\n\n3. HAMBATAN\nPengertian: Usaha yang berasal dari dalam diri sendiri (internal) yang bersifat melemahkan dan menghalangi secara tidak konsepsional pelaksanaan nilai-nilai Pancasila.\nContoh: Rendahnya pemahaman masyarakat tentang Pancasila, kurangnya keteladanan pemimpin, dan korupsi yang merusak kepercayaan terhadap sistem.\n\n4. GANGGUAN\nPengertian: Usaha yang berasal dari luar yang bersifat melemahkan dan menghalangi secara tidak konsepsional pelaksanaan Pancasila.\nContoh: Provokasi dari pihak asing melalui media sosial yang menyebarkan ujaran kebencian dan memecah belah persatuan bangsa.",
    },
    {
        "no": 2, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Ancaman terhadap Pancasila di Era Digital dan Globalisasi",
        "cp": "Peserta didik mampu menganalisis ancaman terhadap Pancasila yang muncul di era digital dan globalisasi",
        "ipk": "Menganalisis bentuk-bentuk ancaman terhadap Pancasila yang muncul akibat globalisasi dan kemajuan digital",
        "indikator": "Diberikan gambaran perkembangan era digital, peserta didik dapat menganalisis ancaman-ancaman terhadap Pancasila yang muncul serta upaya menghadapinya",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Era digital dan globalisasi membawa perubahan besar dalam kehidupan masyarakat Indonesia. Di satu sisi, perkembangan teknologi memberikan manfaat yang luar biasa dalam berbagai bidang. Namun di sisi lain, era digital juga membawa berbagai ancaman terhadap nilai-nilai Pancasila.\n\nMedia sosial yang kini menjadi bagian tak terpisahkan dari kehidupan masyarakat Indonesia ternyata juga menjadi sarana penyebaran paham-paham yang bertentangan dengan Pancasila. Radikalisme, separatisme, hoaks, dan ujaran kebencian dengan mudah menyebar dan mempengaruhi pemikiran masyarakat, terutama generasi muda.\n\nDi sisi lain, penetrasi budaya asing melalui platform digital juga mengancam kelestarian nilai-nilai budaya dan moral bangsa yang berlandaskan Pancasila. Generasi muda sebagai pengguna terbesar media digital perlu memiliki literasi digital yang baik agar tidak mudah terpengaruh.",
        "pertanyaan": "Jelaskan minimal 3 bentuk ancaman terhadap ideologi Pancasila yang muncul di era digital dan globalisasi, serta uraikan upaya yang dapat dilakukan untuk menghadapi setiap ancaman tersebut!",
        "jawaban": "Bentuk Ancaman Pancasila di Era Digital dan Upaya Menghadapinya:\n\n1. Penyebaran Radikalisme dan Ekstremisme melalui Media Digital\nAncaman: Kelompok radikal memanfaatkan media sosial dan platform digital untuk merekrut anggota, menyebarkan ideologi ekstremis, dan mengajak masyarakat menentang Pancasila.\nUpaya: Pemerintah dan masyarakat aktif memblokir konten radikal, meningkatkan literasi digital, serta memperkuat pendidikan karakter berbasis Pancasila di sekolah dan keluarga.\n\n2. Penyebaran Hoaks dan Ujaran Kebencian\nAncaman: Informasi palsu yang disebarkan secara masif dapat memecah belah masyarakat, menimbulkan konflik SARA, dan melemahkan kepercayaan terhadap negara dan Pancasila.\nUpaya: Masyarakat perlu menerapkan prinsip tabayyun (verifikasi) sebelum menyebarkan informasi, melaporkan konten hoaks ke pihak berwenang, dan mendukung lembaga fact-checking.\n\n3. Penetrasi Budaya dan Nilai-Nilai Asing yang Bertentangan dengan Pancasila\nAncaman: Konten digital yang mempromosikan gaya hidup individualistis, hedonisme, dan bebas nilai mengikis nilai-nilai gotong royong, religiositas, dan kemanusiaan yang terkandung dalam Pancasila.\nUpaya: Orang tua dan sekolah berperan aktif mendampingi penggunaan media digital, pemerintah mendorong konten digital bermuatan nilai-nilai Pancasila, dan generasi muda memperkuat identitas budaya bangsa.",
    },
    {
        "no": 3, "elemen": "ATHG terhadap Ideologi Pancasila",
        "materi": "Upaya Bela Negara dalam Menghadapi ATHG terhadap Pancasila",
        "cp": "Peserta didik mampu merumuskan upaya bela negara yang dapat dilakukan untuk menghadapi ATHG terhadap Pancasila",
        "ipk": "Merancang upaya bela negara yang relevan untuk menghadapi berbagai bentuk ATHG terhadap Pancasila",
        "indikator": "Diberikan berbagai bentuk ATHG, peserta didik dapat merumuskan upaya bela negara yang tepat dan relevan untuk menghadapinya",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Bela negara merupakan hak dan kewajiban setiap warga negara Indonesia sebagaimana diamanatkan dalam UUD NRI 1945 Pasal 27 ayat (3) dan Pasal 30. Namun, konsep bela negara tidak hanya terbatas pada pertahanan militer semata, melainkan mencakup segala upaya warga negara untuk mempertahankan eksistensi dan kedaulatan negara, termasuk mempertahankan ideologi Pancasila.\n\nMenghadapi berbagai ATHG terhadap Pancasila di era modern ini memerlukan partisipasi aktif dari seluruh lapisan masyarakat. Bela negara dalam konteks ideologi dapat dilakukan dalam berbagai bentuk yang disesuaikan dengan kemampuan dan profesi masing-masing warga negara.\n\nSebagai pelajar, kontribusi terbaik yang dapat diberikan adalah dengan menjadi generasi yang cerdas, berkarakter Pancasila, dan mampu membentengi diri dari berbagai pengaruh negatif yang dapat melemahkan ideologi bangsa.",
        "pertanyaan": "Jelaskan konsep bela negara dalam konteks menghadapi ATHG terhadap Pancasila, dan uraikan minimal 5 bentuk upaya bela negara yang dapat dilakukan oleh seorang pelajar SMA dalam kehidupan sehari-hari!",
        "jawaban": "Konsep Bela Negara dalam Konteks ATHG terhadap Pancasila:\nBela negara adalah sikap dan perilaku warga negara yang dijiwai kecintaan terhadap NKRI berdasarkan Pancasila dan UUD 1945, dalam menjamin kelangsungan hidup bangsa dan negara. Dalam konteks ATHG terhadap Pancasila, bela negara berarti setiap upaya untuk mempertahankan, mengamalkan, dan menyebarluaskan nilai-nilai Pancasila.\n\nBentuk Upaya Bela Negara bagi Pelajar:\n1. Berprestasi di bidang akademik dan non-akademik sebagai kontribusi nyata bagi kemajuan bangsa serta membuktikan bahwa generasi Pancasila adalah generasi unggul.\n\n2. Mempelajari dan mengamalkan nilai-nilai Pancasila dalam kehidupan sehari-hari, seperti sikap toleran, gotong royong, jujur, dan adil.\n\n3. Aktif melawan penyebaran hoaks dan radikalisme di media sosial dengan tidak menyebarkan informasi yang belum terverifikasi dan melaporkan konten berbahaya.\n\n4. Menjaga persatuan dan kerukunan antar sesama warga negara dengan tidak membeda-bedakan teman berdasarkan SARA (Suku, Agama, Ras, dan Antargolongan).\n\n5. Membangun kesadaran bela negara di lingkungan sekitar dengan menjadi teladan dalam mematuhi aturan, menghormati simbol-simbol negara, dan mencintai produk dalam negeri.",
    },
    {
        "no": 4, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Perbandingan Bentuk Negara Kesatuan dan Negara Federal",
        "cp": "Peserta didik mampu membandingkan bentuk negara kesatuan dengan negara federal beserta kelebihan dan kekurangannya",
        "ipk": "Menganalisis perbedaan, kelebihan, dan kekurangan bentuk negara kesatuan dan negara federal",
        "indikator": "Diberikan informasi tentang dua bentuk negara, peserta didik dapat membandingkan ciri-ciri, kelebihan, dan kekurangan negara kesatuan dan negara federal",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Dalam ilmu negara, dikenal dua bentuk utama negara berdasarkan susunan pemerintahannya, yaitu negara kesatuan (unitaris) dan negara federal (federasi). Setiap bentuk negara memiliki karakteristik, kelebihan, dan kekurangan masing-masing.\n\nIndonesia yang terdiri dari ribuan pulau, ratusan suku bangsa, dan keragaman budaya yang sangat kaya memilih bentuk negara kesatuan sejak kemerdekaan. Pilihan ini didasarkan pada berbagai pertimbangan historis, sosiologis, dan politis yang matang.\n\nSementara itu, beberapa negara besar seperti Amerika Serikat, Jerman, Australia, dan India memilih bentuk negara federal untuk mengakomodasi keberagaman dan kepentingan daerah-daerahnya. Masing-masing bentuk negara ini memiliki implikasi yang berbeda dalam pengelolaan pemerintahan.",
        "pertanyaan": "Bandingkan bentuk negara KESATUAN dan FEDERAL dengan menjelaskan: (a) pengertian masing-masing, (b) minimal 3 ciri/perbedaan, (c) kelebihan dan kekurangan, serta (d) contoh negara yang menganutnya!",
        "jawaban": "Perbandingan Negara Kesatuan dan Negara Federal:\n\n(a) Pengertian:\n- Negara Kesatuan: Negara yang hanya memiliki satu pemerintah pusat yang memegang kedaulatan penuh, daerah-daerah hanya mendapatkan otonomi dari pusat.\n- Negara Federal: Negara yang terdiri dari beberapa negara bagian yang masing-masing memiliki konstitusi dan pemerintahan sendiri, namun tetap tergabung dalam satu kedaulatan bersama.\n\n(b) Perbedaan:\n1. Kedaulatan: Kesatuan - kedaulatan di pusat; Federal - dibagi antara pusat dan negara bagian\n2. Konstitusi: Kesatuan - satu konstitusi; Federal - ada konstitusi pusat dan konstitusi negara bagian\n3. Kewenangan: Kesatuan - daerah mendapat wewenang dari pusat; Federal - negara bagian punya wewenang sendiri\n\n(c) Kelebihan dan Kekurangan:\nNegara Kesatuan: (+) Pemerintahan lebih sederhana dan terintegrasi, kebijakan nasional seragam (-) Kurang fleksibel untuk daerah yang beragam, risiko sentralisasi kekuasaan\nNegara Federal: (+) Lebih fleksibel mengakomodasi keberagaman daerah, mendorong inovasi lokal (-) Birokrasi lebih kompleks, potensi konflik antara pusat dan negara bagian\n\n(d) Contoh Negara:\n- Negara Kesatuan: Indonesia, Perancis, Jepang, Korea Selatan\n- Negara Federal: Amerika Serikat, Jerman, Australia, Malaysia, India",
    },
    {
        "no": 5, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Perbandingan Sistem Pemerintahan Presidensial dan Parlementer",
        "cp": "Peserta didik mampu membandingkan sistem pemerintahan presidensial dan parlementer beserta penerapannya",
        "ipk": "Menganalisis perbedaan, kelebihan, dan kekurangan sistem pemerintahan presidensial dan parlementer",
        "indikator": "Diberikan informasi tentang dua sistem pemerintahan, peserta didik dapat membandingkan ciri, kelebihan, kekurangan presidensial dan parlementer",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Sistem pemerintahan suatu negara menentukan bagaimana kekuasaan eksekutif, legislatif, dan yudikatif diatur dan saling berinteraksi. Secara umum, terdapat dua sistem pemerintahan yang banyak diterapkan di dunia, yaitu sistem presidensial dan sistem parlementer.\n\nIndonesia pada awal kemerdekaan pernah menerapkan sistem parlementer (pada masa UUDS 1950), namun kemudian kembali ke sistem presidensial sesuai dengan UUD 1945. Perjalanan sejarah ini menunjukkan bahwa pemilihan sistem pemerintahan sangat berpengaruh terhadap stabilitas politik dan efektivitas penyelenggaraan pemerintahan.\n\nMemahami perbedaan antara kedua sistem ini penting bagi setiap warga negara agar dapat memahami bagaimana negara mereka dikelola dan apa saja hak serta kewajiban mereka dalam sistem tersebut.",
        "pertanyaan": "Bandingkan sistem pemerintahan PRESIDENSIAL dan PARLEMENTER dengan menjelaskan: (a) pengertian, (b) minimal 4 ciri/perbedaan, (c) kelebihan dan kekurangan masing-masing, dan (d) contoh negara yang menganutnya!",
        "jawaban": "Perbandingan Sistem Pemerintahan Presidensial dan Parlementer:\n\n(a) Pengertian:\n- Presidensial: Sistem pemerintahan di mana kekuasaan eksekutif dipegang oleh presiden yang dipilih langsung oleh rakyat dan tidak bertanggung jawab kepada parlemen.\n- Parlementer: Sistem pemerintahan di mana kekuasaan eksekutif dipegang oleh perdana menteri yang bertanggung jawab kepada parlemen.\n\n(b) Perbedaan:\n1. Kepala Pemerintahan: Presidensial = Presiden; Parlementer = Perdana Menteri\n2. Pemilihan Eksekutif: Presidensial = dipilih rakyat langsung; Parlementer = dipilih/ditunjuk oleh mayoritas parlemen\n3. Pertanggungjawaban: Presidensial = presiden bertanggung jawab kepada rakyat; Parlementer = PM bertanggung jawab kepada parlemen\n4. Masa Jabatan: Presidensial = tetap sesuai konstitusi; Parlementer = dapat dijatuhkan melalui mosi tidak percaya\n\n(c) Kelebihan dan Kekurangan:\nPresidensial: (+) Pemerintahan stabil, masa jabatan pasti; (-) Potensi konflik antara eksekutif dan legislatif\nParlementer: (+) Lebih fleksibel dan responsif; (-) Pemerintahan kurang stabil, sering terjadi pergantian PM\n\n(d) Contoh Negara:\n- Presidensial: Indonesia, Amerika Serikat, Brasil, Meksiko\n- Parlementer: Inggris, Australia, Kanada, India, Malaysia",
    },
    {
        "no": 6, "elemen": "Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
        "materi": "Sistem Pemerintahan Indonesia Berdasarkan UUD NRI 1945",
        "cp": "Peserta didik mampu menjelaskan sistem pemerintahan Indonesia berdasarkan UUD NRI 1945 secara komprehensif",
        "ipk": "Mendeskripsikan sistem pemerintahan Indonesia berdasarkan UUD NRI 1945 setelah amandemen",
        "indikator": "Diberikan pertanyaan tentang sistem pemerintahan Indonesia, peserta didik dapat mendeskripsikan sistem pemerintahan Indonesia pasca amandemen UUD 1945 secara lengkap",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Setelah Reformasi 1998, Indonesia melakukan amandemen UUD 1945 sebanyak empat kali (1999-2002). Amandemen ini membawa perubahan signifikan terhadap sistem pemerintahan Indonesia, terutama dalam hal hubungan antara lembaga-lembaga negara.\n\nSalah satu perubahan terpenting adalah penguatan sistem presidensial murni di Indonesia. Presiden tidak lagi dipilih oleh MPR, melainkan dipilih langsung oleh rakyat melalui pemilihan umum. Selain itu, kekuasaan MPR juga dibatasi, dan lembaga-lembaga negara baru seperti Mahkamah Konstitusi (MK) dan Komisi Yudisial (KY) dibentuk untuk memperkuat sistem checks and balances.\n\nPerubahan-perubahan ini bertujuan untuk mewujudkan pemerintahan yang demokratis, akuntabel, dan bertanggung jawab kepada rakyat sesuai dengan amanat Pancasila dan UUD NRI 1945.",
        "pertanyaan": "Jelaskan sistem pemerintahan Indonesia pasca amandemen UUD NRI 1945! Uraikan: (a) bentuk negara dan pemerintahan, (b) prinsip-prinsip yang dianut, (c) lembaga-lembaga negara utama beserta fungsinya, dan (d) mekanisme checks and balances!",
        "jawaban": "Sistem Pemerintahan Indonesia Pasca Amandemen UUD NRI 1945:\n\n(a) Bentuk Negara dan Pemerintahan:\n- Bentuk Negara: Kesatuan (NKRI)\n- Bentuk Pemerintahan: Republik\n- Sistem Pemerintahan: Presidensial murni\n- Dasar hukum: UUD NRI 1945\n\n(b) Prinsip-prinsip yang Dianut:\n1. Kedaulatan rakyat (Pasal 1 ayat 2)\n2. Negara hukum/rechtsstaat (Pasal 1 ayat 3)\n3. Pemisahan/pembagian kekuasaan (Trias Politica)\n4. Checks and balances antar lembaga negara\n5. Demokrasi konstitusional\n\n(c) Lembaga-lembaga Negara Utama:\n- MPR: Mengubah UUD, melantik Presiden/Wakil Presiden\n- DPR: Legislatif, membuat UU, mengawasi pemerintah\n- DPD: Mewakili kepentingan daerah\n- Presiden: Kepala negara sekaligus kepala pemerintahan\n- MA: Kekuasaan kehakiman umum\n- MK: Menguji UU terhadap UUD, memutus sengketa pemilu\n- BPK: Mengawasi keuangan negara\n\n(d) Mekanisme Checks and Balances:\nDPR mengawasi pemerintah (hak interpelasi, angket, menyatakan pendapat); MK dapat membatalkan UU yang bertentangan UUD; BPK mengaudit keuangan lembaga negara; Presiden bisa mengajukan RUU namun perlu persetujuan DPR.",
    },
    {
        "no": 7, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Hak dan Kewajiban Warga Negara dalam Sistem Pemerintahan RI",
        "cp": "Peserta didik mampu menganalisis hak dan kewajiban warga negara dalam sistem pemerintahan RI berdasarkan UUD 1945",
        "ipk": "Menganalisis hak dan kewajiban warga negara Indonesia dalam pelaksanaan sistem pemerintahan berdasarkan UUD NRI 1945",
        "indikator": "Diberikan pertanyaan tentang hak dan kewajiban warga negara, peserta didik dapat menganalisis hak dan kewajiban warga negara berdasarkan UUD NRI 1945",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Hubungan antara negara dan warga negara dalam sistem pemerintahan Republik Indonesia diatur secara jelas dalam UUD NRI 1945. Warga negara memiliki hak-hak yang harus dijamin pemenuhannya oleh negara, sekaligus memiliki kewajiban yang harus dilaksanakan sebagai bagian dari kehidupan bernegara.\n\nKeseimbangan antara hak dan kewajiban ini merupakan pilar penting dalam sistem demokrasi. Warga negara yang hanya menuntut hak tanpa mau melaksanakan kewajiban, atau sebaliknya negara yang hanya menuntut kewajiban tanpa memenuhi hak warga negara, akan menciptakan ketidakseimbangan yang merusak sistem pemerintahan.\n\nSebagai generasi penerus bangsa, pelajar perlu memahami hak dan kewajibannya sebagai warga negara agar dapat berpartisipasi secara aktif dan bertanggung jawab dalam kehidupan berbangsa dan bernegara.",
        "pertanyaan": "Jelaskan minimal 5 hak dan 5 kewajiban warga negara Indonesia yang diatur dalam UUD NRI 1945! Mengapa keseimbangan antara hak dan kewajiban penting dalam sistem pemerintahan RI?",
        "jawaban": "Hak Warga Negara berdasarkan UUD NRI 1945:\n1. Hak atas perlindungan hukum yang sama (Pasal 27 ayat 1)\n2. Hak atas pekerjaan dan penghidupan yang layak (Pasal 27 ayat 2)\n3. Hak memperoleh pendidikan (Pasal 31 ayat 1)\n4. Hak atas kebebasan berserikat, berkumpul, dan mengeluarkan pendapat (Pasal 28)\n5. Hak atas jaminan sosial (Pasal 34)\n6. Hak memilih dan dipilih dalam pemilihan umum\n\nKewajiban Warga Negara berdasarkan UUD NRI 1945:\n1. Menjunjung tinggi hukum dan pemerintahan (Pasal 27 ayat 1)\n2. Ikut serta dalam upaya pembelaan negara (Pasal 27 ayat 3)\n3. Menghormati HAM orang lain (Pasal 28J ayat 1)\n4. Tunduk pada pembatasan HAM yang ditetapkan UU (Pasal 28J ayat 2)\n5. Ikut serta dalam pertahanan dan keamanan negara (Pasal 30)\n6. Membayar pajak sesuai undang-undang\n\nPentingnya Keseimbangan Hak dan Kewajiban:\nKeseimbangan hak dan kewajiban penting karena:\n- Menciptakan sistem pemerintahan yang adil dan harmonis\n- Memastikan setiap warga negara berkontribusi sekaligus mendapatkan perlindungan\n- Memperkuat fondasi demokrasi yang sehat\n- Mencegah penyalahgunaan hak yang merugikan orang lain",
    },
    {
        "no": 8, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Partisipasi Politik Warga Negara dalam Demokrasi",
        "cp": "Peserta didik mampu menganalisis bentuk-bentuk partisipasi politik warga negara dalam sistem demokrasi Indonesia",
        "ipk": "Menganalisis bentuk partisipasi politik yang efektif dan bertanggung jawab dalam sistem pemerintahan RI",
        "indikator": "Disajikan situasi demokrasi Indonesia, peserta didik dapat menganalisis bentuk partisipasi politik yang efektif dan bertanggung jawab",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Demokrasi yang sehat membutuhkan partisipasi aktif dari seluruh warga negaranya. Dalam sistem pemerintahan Republik Indonesia, warga negara memiliki banyak saluran untuk berpartisipasi dalam kehidupan politik dan pemerintahan, mulai dari keikutsertaan dalam Pemilu, keanggotaan partai politik, hingga keaktifan dalam organisasi masyarakat sipil.\n\nNamun, partisipasi politik yang bermakna tidak hanya sebatas pada keikutsertaan secara formal. Partisipasi yang berkualitas adalah partisipasi yang didasari pemahaman yang baik tentang sistem pemerintahan, dilakukan secara sadar dan bertanggung jawab, serta bertujuan untuk kepentingan bersama bukan sekadar kepentingan golongan.\n\nDi era digital ini, muncul bentuk-bentuk baru partisipasi politik seperti petisi online, kampanye digital, dan advokasi melalui media sosial. Bentuk-bentuk baru ini membuka peluang partisipasi yang lebih luas, namun juga membawa tantangan tersendiri terkait akurasi informasi dan tanggung jawab dalam berekspresi.",
        "pertanyaan": "Jelaskan pentingnya partisipasi politik warga negara dalam sistem pemerintahan RI! Uraikan minimal 4 bentuk partisipasi politik yang dapat dilakukan warga negara, serta jelaskan prinsip-prinsip yang harus dijunjung dalam berpartisipasi politik!",
        "jawaban": "Pentingnya Partisipasi Politik:\nPartisipasi politik warga negara merupakan inti dari sistem demokrasi. Tanpa partisipasi aktif dari rakyat, pemerintahan tidak akan dapat berjalan sesuai dengan kehendak rakyat, akuntabilitas pejabat publik sulit terwujud, dan kepentingan rakyat tidak terwakili dengan baik.\n\nBentuk-bentuk Partisipasi Politik:\n1. Menggunakan Hak Pilih dalam Pemilihan Umum - memilih presiden, kepala daerah, dan anggota legislatif yang kompeten dan berintegritas.\n\n2. Bergabung dengan Partai Politik atau Organisasi Kemasyarakatan - berkontribusi dalam perumusan kebijakan dan pengawasan pemerintah melalui jalur yang terorganisir.\n\n3. Menyampaikan Aspirasi kepada Wakil Rakyat - memanfaatkan mekanisme reses anggota DPR/DPRD, forum musrenbang, atau pertemuan warga untuk menyampaikan kebutuhan dan aspirasi.\n\n4. Aktif Mengawasi Jalannya Pemerintahan - melaporkan dugaan korupsi atau penyimpangan kepada KPK, Ombudsman, atau lembaga pengawas lainnya.\n\nPrinsip-prinsip dalam Berpartisipasi Politik:\n1. Konstitusional: Berpartisipasi melalui jalur yang sah sesuai hukum\n2. Bertanggung jawab: Memahami konsekuensi dari setiap bentuk partisipasi\n3. Berdasarkan informasi yang benar: Tidak terpengaruh hoaks atau provokasi\n4. Berorientasi kepentingan umum: Mendahulukan kepentingan bangsa di atas kepentingan golongan",
    },
    {
        "no": 9, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Sikap Kritis dan Konstruktif terhadap Kebijakan Pemerintah",
        "cp": "Peserta didik mampu menunjukkan sikap kritis dan konstruktif terhadap pelaksanaan sistem pemerintahan RI",
        "ipk": "Menerapkan sikap kritis dan konstruktif yang sesuai dengan nilai demokrasi dalam merespons kebijakan pemerintah",
        "indikator": "Diberikan situasi kebijakan pemerintah yang kontroversial, peserta didik dapat menjelaskan sikap kritis yang konstruktif dan bertanggung jawab",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Dalam sistem demokrasi, warga negara memiliki hak untuk mengkritisi kebijakan pemerintah. Kritik yang konstruktif merupakan bagian penting dari mekanisme kontrol terhadap kekuasaan dan menjadi salah satu pilar utama dalam pemerintahan yang demokratis.\n\nNamun, tidak semua bentuk kritik dapat dibenarkan dalam negara hukum. Kritik yang disampaikan dengan cara-cara yang melanggar hukum, menggunakan informasi palsu, atau bertujuan untuk menghancurkan sistem pemerintahan yang sah, bukanlah bentuk kritik yang konstruktif.\n\nWarga negara yang cerdas dan bertanggung jawab mampu membedakan antara kritik yang membangun dengan tindakan yang merusak tatanan bernegara. Mereka juga memahami bahwa dalam menyampaikan aspirasi, ada mekanisme-mekanisme yang telah disediakan oleh sistem demokrasi untuk menjamin suara rakyat dapat didengar.",
        "pertanyaan": "Jelaskan apa yang dimaksud dengan sikap kritis dan konstruktif terhadap pemerintahan! Uraikan: (a) perbedaan kritik konstruktif dengan tindakan destruktif, (b) saluran resmi yang tersedia bagi warga negara untuk menyampaikan kritik, dan (c) batasan dalam menyampaikan kritik berdasarkan hukum yang berlaku!",
        "jawaban": "Pengertian Sikap Kritis dan Konstruktif:\nSikap kritis adalah kemampuan untuk menilai, menganalisis, dan mengevaluasi kebijakan atau tindakan pemerintah secara objektif berdasarkan fakta. Sikap konstruktif berarti kritik tersebut disertai dengan solusi atau alternatif yang membangun, bukan sekadar menjatuhkan.\n\n(a) Perbedaan Kritik Konstruktif dengan Tindakan Destruktif:\nKritik Konstruktif:\n- Berdasarkan fakta dan data yang valid\n- Disampaikan melalui saluran yang sah\n- Disertai saran/solusi perbaikan\n- Bertujuan memperbaiki keadaan\n\nTindakan Destruktif:\n- Berdasarkan hoaks atau informasi yang dimanipulasi\n- Dilakukan melalui cara-cara yang melanggar hukum\n- Bertujuan menghancurkan sistem, bukan memperbaikinya\n- Menghasut kebencian dan kekerasan\n\n(b) Saluran Resmi Penyampaian Kritik:\n- Menyampaikan aspirasi kepada anggota DPR/DPRD pada masa reses\n- Forum musrenbang di tingkat desa/kelurahan, kecamatan, hingga nasional\n- Melaporkan ke lembaga pengawas: KPK (korupsi), Ombudsman (mal administrasi)\n- Menyampaikan petisi atau surat terbuka kepada pejabat yang berwenang\n- Menyampaikan pendapat melalui media massa yang bertanggung jawab\n\n(c) Batasan dalam Menyampaikan Kritik:\n- UUD 1945 Pasal 28 menjamin kebebasan berpendapat\n- UU ITE mengatur pembatasan dalam berekspresi di media digital\n- Kritik tidak boleh memuat SARA, pencemaran nama baik, atau ancaman\n- Unjuk rasa harus mendapat izin dari kepolisian sesuai UU No. 9 Tahun 1998",
    },
    {
        "no": 10, "elemen": "Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan Republik Indonesia",
        "materi": "Peran Generasi Muda dalam Mendukung Sistem Pemerintahan RI",
        "cp": "Peserta didik mampu merumuskan peran aktif generasi muda dalam mendukung pelaksanaan sistem pemerintahan RI yang baik",
        "ipk": "Merancang kontribusi nyata generasi muda dalam mendukung terwujudnya good governance di Indonesia",
        "indikator": "Diberikan tantangan dalam penyelenggaraan pemerintahan, peserta didik dapat merumuskan peran aktif generasi muda dalam mendukung sistem pemerintahan RI yang baik",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Generasi muda Indonesia merupakan aset bangsa yang sangat berharga. Mereka adalah penerus estafet kepemimpinan bangsa yang kelak akan menentukan arah perjalanan negara ini. Oleh karena itu, peran aktif generasi muda dalam mendukung sistem pemerintahan yang baik sangatlah penting.\n\nDi era reformasi ini, Indonesia terus berupaya mewujudkan tata kelola pemerintahan yang baik (good governance) yang ditandai dengan transparansi, akuntabilitas, partisipasi, dan supremasi hukum. Namun, upaya ini tidak dapat dilakukan oleh pemerintah saja tanpa dukungan aktif dari seluruh warga negara, termasuk generasi muda.\n\nGenerasi muda yang melek teknologi, kritis, dan bersemangat memiliki potensi besar untuk menjadi agen perubahan dalam mewujudkan Indonesia yang lebih baik. Dengan pendidikan dan karakter yang kuat berlandaskan Pancasila, mereka dapat menjadi garda terdepan dalam membangun Indonesia.",
        "pertanyaan": "Jelaskan apa yang dimaksud dengan good governance! Uraikan minimal 5 peran nyata yang dapat dilakukan oleh generasi muda pelajar SMA dalam mendukung terwujudnya good governance di Indonesia dalam kehidupan sehari-hari!",
        "jawaban": "Pengertian Good Governance:\nGood governance (tata kelola pemerintahan yang baik) adalah penyelenggaraan pemerintahan yang transparan, akuntabel, partisipatif, efektif, efisien, dan sesuai dengan hukum yang berlaku. Good governance melibatkan tiga pilar: pemerintah, sektor swasta, dan masyarakat sipil yang bekerja sama secara harmonis.\n\nPrinsip-prinsip Good Governance: Transparansi, Akuntabilitas, Partisipasi, Supremasi Hukum, Efektivitas dan Efisiensi, Keadilan dan Inklusivitas.\n\nPeran Nyata Generasi Muda Pelajar dalam Mendukung Good Governance:\n\n1. Menerapkan Nilai Anti Korupsi dalam Kehidupan Sehari-hari\nJujur dalam ulangan, tidak mencontek, tidak menerima hadiah yang bermuatan sogok, dan terbiasa hidup dengan prinsip integritas sejak dini.\n\n2. Memanfaatkan Teknologi secara Positif dan Bertanggung Jawab\nMenyebarkan informasi yang benar dan terverifikasi, melawan hoaks dan ujaran kebencian, serta memanfaatkan media sosial untuk edukasi dan kampanye positif.\n\n3. Aktif dalam Kegiatan Demokrasi di Sekolah\nBerpartisipasi dalam pemilihan OSIS secara jujur dan demokratis, menyampaikan aspirasi melalui forum yang tersedia, dan menghargai hasil keputusan bersama.\n\n4. Membangun Solidaritas dan Kepedulian Sosial\nAktif dalam kegiatan sosial kemasyarakatan, membantu sesama tanpa memandang perbedaan, dan memperkuat persatuan antar warga sekolah.\n\n5. Terus Belajar dan Meningkatkan Kompetensi\nBerprestasi untuk mempersiapkan diri menjadi pemimpin masa depan yang kompeten, berintegritas, dan berjiwa Pancasila demi Indonesia yang lebih baik.",
    },
]

def rebuild_cell(cell, text):
    tc = cell._tc
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    lines = text.split('\n') if text else ['']
    for line in lines:
        p = OxmlElement('w:p')
        if line:
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = line
            if line.startswith(' ') or line.endswith(' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.append(t)
            p.append(r)
        tc.append(p)

def fix_header_cell(cell):
    for para in cell.paragraphs:
        for run in para.runs:
            t = run.text
            if '..............di isi' in t:
                run.text = t.replace('..............di isi', NAMA_PENULIS)
            elif '.............di isi' in t:
                run.text = t.replace('.............di isi', f'{MAPEL} / Kelas {KELAS}')
            elif '...............' in t and len(t.strip('.')) == 0:
                run.text = NAMA_PENULIS

def fill_pg_table(table, q):
    fix_header_cell(table.rows[0].cells[0])
    # Referensi
    for r in [1, 2]:
        for c in [1, 2, 3, 4]:
            try:
                cell = table.rows[r].cells[c]
                if id(cell._tc) != id(table.rows[0].cells[0]._tc):
                    rebuild_cell(cell, f"Buku Acuan/Referensi:\n{REF}")
            except:
                pass
    # Metadata kiri
    try: rebuild_cell(table.rows[2].cells[0], q['elemen'])
    except: pass
    try: rebuild_cell(table.rows[5].cells[0], q['cp'])
    except: pass
    try: rebuild_cell(table.rows[6].cells[0], f"Materi: {q['materi']}")
    except: pass
    try:
        if id(table.rows[6].cells[0]._tc) != id(table.rows[7].cells[0]._tc):
            rebuild_cell(table.rows[7].cells[0], f"Kelas: {KELAS} / Semester: Genap")
    except: pass
    try: rebuild_cell(table.rows[9].cells[0], f"Indikator Pencapaian Kompetensi:\n{q['ipk']}")
    except: pass
    try:
        if id(table.rows[10].cells[0]._tc) != id(table.rows[9].cells[0]._tc):
            rebuild_cell(table.rows[10].cells[0], q['ipk'])
    except: pass
    try:
        tingkat = q.get('tingkat','Sedang')
        rebuild_cell(table.rows[11].cells[0],
            f"Indikator Soal / Indikator Asessmen:\n{q['indikator']}\n\n"
            f"Tingkat Kesukaran:\n"
            f"{'[v]' if tingkat=='Rendah' else '[ ]'} Rendah\n"
            f"{'[v]' if tingkat=='Sedang' else '[ ]'} Sedang\n"
            f"{'[v]' if tingkat=='Sulit' else '[ ]'} Sulit")
    except: pass
    # Nomor soal
    try: rebuild_cell(table.rows[4].cells[2], f"No. Soal\n{q['no']}")
    except: pass
    # Nilai
    try:
        c8 = table.rows[8].cells[2]
        seen = {id(table.rows[r].cells[2]._tc) for r in range(4,8)}
        if id(c8._tc) not in seen:
            rebuild_cell(c8, str(q.get('nilai','5')))
    except: pass
    # Kunci Jawaban
    seen_k = set()
    first = True
    for r in [11,12,13]:
        try:
            cell = table.rows[r].cells[2]
            cid = id(cell._tc)
            if cid not in seen_k:
                seen_k.add(cid)
                if first:
                    rebuild_cell(cell, f"KUNCI\nJAWABAN\n\n{q['jawaban']}")
                    first = False
                else:
                    rebuild_cell(cell, "")
        except: pass
    # Soal
    soal_text = (f"{q['soal']}\n\nA. {q['A']}\n\nB. {q['B']}\n\n"
                 f"C. {q['C']}\n\nD. {q['D']}\n\nE. {q['E']}")
    seen_q = set(); first_q = True
    for r in range(3,14):
        try:
            cell = table.rows[r].cells[3]
            cid = id(cell._tc)
            if cid not in seen_q and 'Ketika Soal' in cell.text:
                seen_q.add(cid)
                if first_q:
                    rebuild_cell(cell, soal_text); first_q = False
                else:
                    rebuild_cell(cell, "")
        except: pass

def fill_essay_table(table, q):
    fix_header_cell(table.rows[0].cells[0])
    for r in [1,2]:
        for c in [1,2,3,4]:
            try:
                cell = table.rows[r].cells[c]
                if id(cell._tc) != id(table.rows[0].cells[0]._tc):
                    rebuild_cell(cell, f"Buku Acuan/Referensi:\n{REF}")
            except: pass
    try: rebuild_cell(table.rows[2].cells[0], q['elemen'])
    except: pass
    try: rebuild_cell(table.rows[5].cells[0], q['cp'])
    except: pass
    try: rebuild_cell(table.rows[6].cells[0], f"Materi: {q['materi']}")
    except: pass
    try:
        if id(table.rows[6].cells[0]._tc) != id(table.rows[7].cells[0]._tc):
            rebuild_cell(table.rows[7].cells[0], f"Kelas: {KELAS} / Semester: Genap")
    except: pass
    try: rebuild_cell(table.rows[9].cells[0], f"Indikator Pencapaian Kompetensi:\n{q['ipk']}")
    except: pass
    try:
        if id(table.rows[10].cells[0]._tc) != id(table.rows[9].cells[0]._tc):
            rebuild_cell(table.rows[10].cells[0], q['ipk'])
    except: pass
    try:
        tingkat = q.get('tingkat','Sedang')
        rebuild_cell(table.rows[11].cells[0],
            f"Indikator Soal / Indikator Asessmen:\n{q['indikator']}\n\n"
            f"Tingkat Kesukaran:\n"
            f"{'[v]' if tingkat=='Rendah' else '[ ]'} Rendah\n"
            f"{'[v]' if tingkat=='Sedang' else '[ ]'} Sedang\n"
            f"{'[v]' if tingkat=='Sulit' else '[ ]'} Sulit")
    except: pass
    try: rebuild_cell(table.rows[4].cells[2], f"No. Soal\n{q['no']}")
    except: pass
    try:
        c8 = table.rows[8].cells[2]
        seen = {id(table.rows[r].cells[2]._tc) for r in range(4,8)}
        if id(c8._tc) not in seen:
            rebuild_cell(c8, f"{q.get('nilai','10')}\n(poin)")
    except: pass
    seen_k = set()
    for r in [11,12,13]:
        try:
            cell = table.rows[r].cells[2]
            cid = id(cell._tc)
            if cid not in seen_k:
                seen_k.add(cid)
                rebuild_cell(cell, "KUNCI\nJAWABAN")
        except: pass
    soal_text = (f"{q['narasi']}\n\nPERTANYAAN:\n{q['pertanyaan']}\n\n"
                 f"---KUNCI JAWABAN---\n\n{q['jawaban']}")
    seen_q = set(); first_q = True
    for r in range(3,14):
        try:
            cell = table.rows[r].cells[3]
            cid = id(cell._tc)
            if cid not in seen_q and ('Ketika Soal' in cell.text or 'Ketik Jawaban' in cell.text):
                seen_q.add(cid)
                if first_q:
                    rebuild_cell(cell, soal_text); first_q = False
                else:
                    rebuild_cell(cell, "")
        except: pass

def fill_kisi_kisi(table):
    pg_groups = [
        (1,"Ancaman Tantangan Hambatan dan Gangguan (ATHG) terhadap Ideologi Pancasila",
         "Peserta didik mampu menganalisis berbagai bentuk ATHG terhadap Pancasila dan upaya menghadapinya",
         "Pengertian, jenis, dan upaya menghadapi ATHG terhadap Pancasila",
         "Menganalisis ATHG terhadap Pancasila dan menentukan upaya penanggulangannya",
         "7","1 - 7","Pilihan Ganda","Sedang - Sulit"),
        (2,"Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
         "Peserta didik mampu membedakan bentuk negara, bentuk pemerintahan, dan sistem pemerintahan",
         "Negara kesatuan, federal, monarki, republik, presidensial, parlementer",
         "Membedakan bentuk negara dan sistem pemerintahan beserta contohnya",
         "7","8 - 14","Pilihan Ganda","Rendah - Sedang"),
        (3,"Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan RI",
         "Peserta didik mampu menunjukkan sikap positif warga negara terhadap sistem pemerintahan RI",
         "Partisipasi, hak-kewajiban, sikap kritis konstruktif, taat hukum",
         "Menerapkan sikap positif warga negara terhadap sistem pemerintahan RI",
         "6","15 - 20","Pilihan Ganda","Rendah - Sulit"),
    ]
    essay_groups = [
        (1,"ATHG terhadap Ideologi Pancasila",
         "Peserta didik mampu menganalisis ATHG terhadap Pancasila dan upaya bela negara",
         "Klasifikasi ATHG, ancaman digital, dan upaya bela negara",
         "Menjelaskan dan menganalisis ATHG terhadap Pancasila beserta upaya menghadapinya",
         "3","1 - 3","Uraian/Essay","Sedang - Sulit"),
        (2,"Bentuk Negara, Bentuk Pemerintahan, dan Sistem Pemerintahan",
         "Peserta didik mampu membandingkan dan menganalisis bentuk negara dan sistem pemerintahan",
         "Negara kesatuan vs federal, presidensial vs parlementer, sistem pemerintahan RI",
         "Membandingkan bentuk negara dan sistem pemerintahan serta menganalisis sistem pemerintahan RI",
         "3","4 - 6","Uraian/Essay","Sedang - Sulit"),
        (3,"Sikap Warga Negara terhadap Pelaksanaan Sistem Pemerintahan RI",
         "Peserta didik mampu menganalisis sikap warga negara dan peran generasi muda",
         "Hak-kewajiban, partisipasi politik, sikap kritis, peran pemuda",
         "Menganalisis hak-kewajiban, partisipasi, dan peran generasi muda dalam sistem pemerintahan",
         "4","7 - 10","Uraian/Essay","Sulit"),
    ]
    for no,elemen,cp,materi,tp,jml,nomor,bentuk,tingkat in pg_groups:
        row = table.rows[no]
        for c,v in enumerate([str(no),elemen,cp,materi,tp,jml,nomor,bentuk,tingkat]):
            try: rebuild_cell(row.cells[c], v)
            except: pass
    for no,elemen,cp,materi,tp,jml,nomor,bentuk,tingkat in essay_groups:
        row = table.rows[no+5]
        for c,v in enumerate([str(no),elemen,cp,materi,tp,jml,nomor,bentuk,tingkat]):
            try: rebuild_cell(row.cells[c], v)
            except: pass
    try:
        row = table.rows[11]
        rebuild_cell(row.cells[0],"TOTAL")
        rebuild_cell(row.cells[5],"30")
        rebuild_cell(row.cells[6],"PG: 1-20\nEssay: 1-10")
        rebuild_cell(row.cells[7],"PG & Essay")
        rebuild_cell(row.cells[8],"Variatif")
    except: pass

def fill_table0(table):
    cell = table.rows[0].cells[0]
    for para in cell.paragraphs:
        for run in para.runs:
            if '...............................' in run.text:
                run.text = run.text.replace('...............................', NAMA_PENULIS)

def main():
    doc = Document('/root/.claude/uploads/19d916f2-492a-4a0c-bfba-2d98a5361dd0/faf676bf-KARTU_SOAL_UasGENAP_2526okeaslikosongan.docx')
    
    print("Menambah 7 tabel PG tambahan...")
    ref_elem = doc.tables[2]._element
    last_pg = doc.tables[14]._element
    parent = last_pg.getparent()
    idx14 = list(parent).index(last_pg)
    for i in range(7):
        parent.insert(idx14+1+i, deepcopy(ref_elem))
    print(f"Total tabel: {len(doc.tables)}")
    
    fill_table0(doc.tables[0])
    fill_kisi_kisi(doc.tables[1])
    
    for i, q in enumerate(pg_data):
        print(f"  PG Soal {q['no']} -> Tabel {i+2}")
        fill_pg_table(doc.tables[i+2], q)
    
    for i, q in enumerate(essay_data):
        print(f"  Essay Soal {q['no']} -> Tabel {22+i}")
        fill_essay_table(doc.tables[22+i], q)
    
    out = '/home/user/daw/KARTU_SOAL_UAS_GENAP_PKN_XI.docx'
    doc.save(out)
    print(f"\nSelesai! Tersimpan: {out}")

main()
