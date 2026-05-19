from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
from docx.oxml import OxmlElement

NAMA_PENULIS = "ROBY PRASETYO ADI"
MAPEL = "Pendidikan Kewarganegaraan (PKN)"
KELAS = "X"
REF = "Buku Siswa PKN Kelas X, Kemendikbud RI; Modul PKN Merdeka Belajar Kelas X; UUD NRI 1945"

pg_data = [
    # ---- MATERI 1: HAK DAN KEWAJIBAN SEBAGAI WARGA SEKOLAH DAN MASYARAKAT (Soal 1-5) ----
    {
        "no": 1, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Pengertian Hak dan Kewajiban Warga Sekolah",
        "cp": "Peserta didik mampu memahami dan membedakan hak dan kewajiban sebagai warga sekolah",
        "ipk": "Membedakan pengertian hak dan kewajiban warga sekolah beserta contohnya",
        "indikator": "Disajikan deskripsi tentang kegiatan di lingkungan sekolah, peserta didik dapat membedakan yang merupakan hak dan yang merupakan kewajiban warga sekolah",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Dalam kehidupan sehari-hari di lingkungan sekolah, setiap warga sekolah memiliki hak dan kewajiban yang harus dipahami dan dilaksanakan. Berikut ini yang merupakan HAK peserta didik di sekolah adalah ...",
        "A": "Mengerjakan tugas yang diberikan oleh guru dengan tepat waktu",
        "B": "Mematuhi tata tertib sekolah yang telah ditetapkan",
        "C": "Mendapatkan pengajaran, bimbingan, dan fasilitas belajar yang memadai dari sekolah",
        "D": "Menjaga kebersihan dan ketertiban lingkungan sekolah",
        "E": "Mengikuti kegiatan upacara bendera setiap hari Senin",
        "jawaban": "C",
    },
    {
        "no": 2, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Kewajiban Warga Sekolah",
        "cp": "Peserta didik mampu mengidentifikasi kewajiban warga sekolah dan melaksanakannya dalam kehidupan nyata",
        "ipk": "Mengidentifikasi kewajiban peserta didik, guru, dan tenaga kependidikan di sekolah",
        "indikator": "Diberikan daftar perilaku di lingkungan sekolah, peserta didik dapat mengidentifikasi yang termasuk kewajiban warga sekolah",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Perhatikan pernyataan-pernyataan berikut:\n1. Mendapatkan nilai yang adil sesuai kemampuan\n2. Mengikuti seluruh kegiatan pembelajaran dengan serius\n3. Memperoleh perlindungan dari kekerasan di sekolah\n4. Menghormati guru dan tenaga kependidikan\n5. Menggunakan fasilitas sekolah dengan bertanggung jawab\nYang termasuk KEWAJIBAN peserta didik di sekolah adalah ...",
        "A": "1, 2, dan 3",
        "B": "2, 4, dan 5",
        "C": "1, 3, dan 5",
        "D": "2, 3, dan 4",
        "E": "1, 2, dan 5",
        "jawaban": "B",
    },
    {
        "no": 3, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Hak dan Kewajiban Warga Masyarakat",
        "cp": "Peserta didik mampu memahami hak dan kewajiban warga dalam kehidupan bermasyarakat",
        "ipk": "Menjelaskan hak dan kewajiban warga dalam kehidupan bermasyarakat berdasarkan norma yang berlaku",
        "indikator": "Disajikan situasi kehidupan bermasyarakat, peserta didik dapat menjelaskan hak dan kewajiban warga masyarakat",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Keseimbangan antara hak dan kewajiban dalam kehidupan bermasyarakat sangat penting untuk menciptakan kehidupan yang harmonis. Seorang warga masyarakat yang melaksanakan kewajibannya membayar iuran kebersihan lingkungan berhak mendapatkan ...",
        "A": "Pujian dari seluruh warga masyarakat sekitar",
        "B": "Layanan kebersihan yang baik dan lingkungan yang bersih dan sehat",
        "C": "Pengecualian dari kewajiban-kewajiban lain di masyarakat",
        "D": "Jabatan sebagai pengurus RT/RW setempat",
        "E": "Kompensasi finansial atas iuran yang telah dibayarkan",
        "jawaban": "B",
    },
    {
        "no": 4, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Pelaksanaan Hak dan Kewajiban secara Seimbang",
        "cp": "Peserta didik mampu menganalisis pentingnya keseimbangan hak dan kewajiban dalam kehidupan sekolah dan masyarakat",
        "ipk": "Menganalisis akibat dari tidak seimbangnya pelaksanaan hak dan kewajiban dalam kehidupan bermasyarakat",
        "indikator": "Diberikan kasus ketidakseimbangan hak dan kewajiban, peserta didik dapat menganalisis dampak dan solusinya",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Seseorang yang selalu menuntut haknya namun tidak mau melaksanakan kewajibannya dalam kehidupan bermasyarakat akan menyebabkan ...",
        "A": "Kehidupan masyarakat menjadi lebih maju dan sejahtera",
        "B": "Terciptanya iklim kompetisi yang sehat di masyarakat",
        "C": "Timbulnya konflik, ketidakharmonisan, dan ketidakadilan di masyarakat",
        "D": "Mendorong orang lain untuk lebih aktif melaksanakan kewajiban",
        "E": "Meningkatnya kesadaran warga masyarakat tentang pentingnya hak",
        "jawaban": "C",
    },
    {
        "no": 5, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Contoh Pelaksanaan Hak dan Kewajiban di Masyarakat",
        "cp": "Peserta didik mampu menerapkan hak dan kewajiban warga masyarakat dalam kehidupan sehari-hari",
        "ipk": "Menerapkan sikap bertanggung jawab dalam melaksanakan hak dan kewajiban sebagai warga masyarakat",
        "indikator": "Disajikan situasi kehidupan sehari-hari, peserta didik dapat menentukan sikap yang mencerminkan pelaksanaan hak dan kewajiban secara bertanggung jawab",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Dalam kehidupan bermasyarakat, sikap yang paling mencerminkan pelaksanaan hak dan kewajiban secara seimbang dan bertanggung jawab adalah ...",
        "A": "Menuntut pemerintah memperbaiki jalan tanpa ikut menjaga kebersihan lingkungan",
        "B": "Mengikuti rapat warga, membayar iuran, menjaga keamanan lingkungan, dan menggunakan hak suara dalam pemilihan RT",
        "C": "Hanya melaksanakan kewajiban tanpa mempedulikan hak yang seharusnya diterima",
        "D": "Memanfaatkan semua fasilitas umum tanpa ikut berkontribusi dalam pemeliharaannya",
        "E": "Mengutamakan kepentingan pribadi di atas kepentingan masyarakat dalam setiap keputusan",
        "jawaban": "B",
    },
    # ---- MATERI 2: SISTEM PERTAHANAN DAN KEAMANAN NEGARA INDONESIA (Soal 6-10) ----
    {
        "no": 6, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Pengertian dan Dasar Hukum Sistem Pertahanan Negara",
        "cp": "Peserta didik mampu memahami sistem pertahanan dan keamanan negara Indonesia berdasarkan UUD NRI 1945",
        "ipk": "Menjelaskan pengertian dan dasar hukum sistem pertahanan dan keamanan negara Indonesia",
        "indikator": "Diberikan pernyataan tentang pertahanan negara, peserta didik dapat menjelaskan pengertian dan dasar hukum sistem pertahanan dan keamanan negara Indonesia",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Sistem pertahanan dan keamanan negara Indonesia diatur dalam UUD NRI 1945. Pasal yang secara khusus mengatur tentang pertahanan dan keamanan negara adalah ...",
        "A": "Pasal 27 ayat (1)",
        "B": "Pasal 28 ayat (1)",
        "C": "Pasal 30 ayat (1) dan (2)",
        "D": "Pasal 31 ayat (1)",
        "E": "Pasal 33 ayat (1)",
        "jawaban": "C",
    },
    {
        "no": 7, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Sistem Pertahanan Rakyat Semesta (Sishankamrata)",
        "cp": "Peserta didik mampu memahami konsep Sistem Pertahanan Rakyat Semesta (Sishankamrata)",
        "ipk": "Menjelaskan konsep dan ciri-ciri Sistem Pertahanan Rakyat Semesta (Sishankamrata) Indonesia",
        "indikator": "Disajikan deskripsi tentang sistem pertahanan Indonesia, peserta didik dapat menjelaskan konsep Sishankamrata dan peran komponen-komponennya",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Indonesia menganut sistem pertahanan yang bersifat semesta, yang dikenal dengan Sistem Pertahanan Rakyat Semesta (Sishankamrata). Ciri utama Sishankamrata adalah ...",
        "A": "Pertahanan negara hanya mengandalkan kekuatan militer profesional TNI",
        "B": "Seluruh rakyat, wilayah, dan sumber daya nasional dilibatkan sebagai kekuatan pertahanan",
        "C": "Keamanan negara diserahkan sepenuhnya kepada Polri tanpa melibatkan rakyat",
        "D": "Pertahanan negara hanya difokuskan pada wilayah perbatasan dan lautan",
        "E": "Kekuatan pertahanan didatangkan dari negara-negara sekutu yang telah bersepakat",
        "jawaban": "B",
    },
    {
        "no": 8, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Komponen Pertahanan Negara",
        "cp": "Peserta didik mampu menganalisis komponen-komponen dalam sistem pertahanan negara Indonesia",
        "ipk": "Membedakan komponen utama, komponen cadangan, dan komponen pendukung dalam sistem pertahanan negara",
        "indikator": "Diberikan daftar unsur pertahanan, peserta didik dapat mengklasifikasikan komponen utama, cadangan, dan pendukung sistem pertahanan negara",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Dalam Undang-Undang Nomor 3 Tahun 2002 tentang Pertahanan Negara, komponen pertahanan negara terdiri dari komponen utama, cadangan, dan pendukung. Yang merupakan KOMPONEN UTAMA pertahanan negara Indonesia adalah ...",
        "A": "Seluruh warga negara yang memiliki kemampuan fisik dan mental",
        "B": "Tentara Nasional Indonesia (TNI)",
        "C": "Organisasi kemasyarakatan yang terlatih bela negara",
        "D": "Sumber daya alam dan sumber daya buatan",
        "E": "Polisi Republik Indonesia (Polri)",
        "jawaban": "B",
    },
    {
        "no": 9, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Peran Warga Negara dalam Pertahanan dan Keamanan",
        "cp": "Peserta didik mampu menganalisis peran warga negara dalam sistem pertahanan dan keamanan negara",
        "ipk": "Menganalisis bentuk partisipasi warga negara dalam mendukung sistem pertahanan dan keamanan negara",
        "indikator": "Diberikan berbagai perilaku warga negara, peserta didik dapat menganalisis bentuk partisipasi yang mendukung sistem pertahanan dan keamanan negara",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Pasal 30 ayat (1) UUD NRI 1945 menyatakan bahwa tiap-tiap warga negara berhak dan wajib ikut serta dalam usaha pertahanan dan keamanan negara. Bentuk partisipasi warga negara (bukan TNI/Polri) yang paling tepat dalam mendukung sistem pertahanan dan keamanan negara adalah ...",
        "A": "Membeli dan menyimpan senjata api untuk melindungi diri dan keluarga",
        "B": "Menolak semua pengaruh asing masuk ke Indonesia",
        "C": "Aktif dalam kegiatan siskamling, melaporkan hal mencurigakan kepada pihak berwajib, dan menjaga kondusivitas lingkungan",
        "D": "Bergabung dengan milisi bersenjata untuk membantu TNI di perbatasan",
        "E": "Mendirikan pos keamanan pribadi di lingkungan tempat tinggal",
        "jawaban": "C",
    },
    {
        "no": 10, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Bela Negara sebagai Kewajiban Warga Negara",
        "cp": "Peserta didik mampu memahami konsep bela negara sebagai kewajiban warga negara Indonesia",
        "ipk": "Menjelaskan konsep dan bentuk-bentuk bela negara yang dapat dilakukan oleh warga negara",
        "indikator": "Diberikan konsep bela negara, peserta didik dapat menjelaskan bentuk-bentuk bela negara yang dapat dilakukan oleh warga negara biasa",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Bela negara merupakan kewajiban setiap warga negara Indonesia. Bagi seorang pelajar SMA, bentuk bela negara yang paling relevan dan dapat dilakukan dalam kehidupan sehari-hari adalah ...",
        "A": "Mendaftar sebagai anggota TNI aktif setelah lulus SMA",
        "B": "Mengikuti latihan militer setiap akhir pekan di instansi pertahanan",
        "C": "Belajar dengan sungguh-sungguh, berprestasi, dan mengamalkan nilai-nilai Pancasila dalam kehidupan",
        "D": "Mengumpulkan dana untuk membeli persenjataan bagi TNI",
        "E": "Menjaga pos keamanan di lingkungan sekolah secara mandiri",
        "jawaban": "C",
    },
    # ---- MATERI 3: PERAN INDONESIA DALAM HUBUNGAN ANTAR NEGARA (Soal 11-15) ----
    {
        "no": 11, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Dasar Hukum Politik Luar Negeri Indonesia",
        "cp": "Peserta didik mampu memahami dasar hukum dan prinsip politik luar negeri Indonesia",
        "ipk": "Menjelaskan dasar hukum dan prinsip bebas aktif dalam politik luar negeri Indonesia",
        "indikator": "Diberikan pertanyaan tentang politik luar negeri Indonesia, peserta didik dapat menjelaskan dasar hukum dan prinsip bebas aktif",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Indonesia menjalankan politik luar negeri yang bebas dan aktif. Prinsip 'bebas aktif' dalam politik luar negeri Indonesia berarti ...",
        "A": "Indonesia bebas bergabung dengan blok manapun dan aktif menyerang negara lain",
        "B": "Indonesia tidak terikat pada blok kekuatan manapun dan aktif berperan dalam perdamaian dunia",
        "C": "Indonesia bebas menerima bantuan asing dan aktif meminta dukungan internasional",
        "D": "Indonesia bebas dari pengaruh PBB dan aktif menjalin perdagangan internasional",
        "E": "Indonesia bebas menentukan kebijakan ekonomi dan aktif bersaing dengan negara lain",
        "jawaban": "B",
    },
    {
        "no": 12, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Peran Indonesia dalam ASEAN",
        "cp": "Peserta didik mampu menganalisis peran aktif Indonesia dalam organisasi regional ASEAN",
        "ipk": "Menjelaskan peran dan kontribusi Indonesia dalam organisasi ASEAN",
        "indikator": "Diberikan informasi tentang ASEAN, peserta didik dapat menjelaskan peran dan kontribusi penting Indonesia dalam ASEAN",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Indonesia merupakan salah satu negara pendiri ASEAN yang memiliki peran strategis dalam organisasi tersebut. Salah satu peran penting Indonesia dalam ASEAN yang menunjukkan kepemimpinan aktif adalah ...",
        "A": "Indonesia selalu memenangkan semua kompetisi olahraga dalam SEA Games",
        "B": "Indonesia sebagai negara dengan jumlah penduduk terbanyak di ASEAN",
        "C": "Indonesia menjadi pemrakarsa berdirinya ASEAN tahun 1967 dan aktif menjadi penengah dalam penyelesaian konflik di kawasan",
        "D": "Indonesia menjadi negara dengan perekonomian terbesar di kawasan Asia Tenggara",
        "E": "Indonesia selalu menduduki posisi Sekretaris Jenderal ASEAN setiap periode",
        "jawaban": "C",
    },
    {
        "no": 13, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Peran Indonesia dalam PBB",
        "cp": "Peserta didik mampu menganalisis peran Indonesia dalam Perserikatan Bangsa-Bangsa (PBB)",
        "ipk": "Menjelaskan kontribusi Indonesia dalam misi perdamaian dan forum-forum PBB",
        "indikator": "Diberikan informasi tentang peran Indonesia di PBB, peserta didik dapat menjelaskan kontribusi Indonesia dalam misi perdamaian dan forum internasional",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Indonesia aktif berperan dalam Perserikatan Bangsa-Bangsa (PBB). Salah satu bentuk kontribusi nyata Indonesia dalam PBB di bidang perdamaian dunia adalah ...",
        "A": "Indonesia selalu menentang semua resolusi yang diajukan oleh PBB",
        "B": "Indonesia mengirimkan pasukan Garuda untuk misi perdamaian PBB (UNPKF) di berbagai negara konflik",
        "C": "Indonesia menjadi anggota tetap Dewan Keamanan PBB yang memiliki hak veto",
        "D": "Indonesia mendanai seluruh kegiatan operasional PBB di kawasan Asia Pasifik",
        "E": "Indonesia menjadi tuan rumah tetap sidang umum PBB setiap tahunnya",
        "jawaban": "B",
    },
    {
        "no": 14, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Kerja Sama Internasional dan Manfaatnya bagi Indonesia",
        "cp": "Peserta didik mampu menganalisis manfaat kerja sama internasional bagi Indonesia",
        "ipk": "Menganalisis manfaat kerja sama internasional bagi pembangunan dan kepentingan nasional Indonesia",
        "indikator": "Diberikan contoh kerja sama internasional, peserta didik dapat menganalisis manfaat kerja sama tersebut bagi Indonesia",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Indonesia aktif menjalin kerja sama dengan berbagai negara dan organisasi internasional. Manfaat utama kerja sama internasional bagi Indonesia adalah ...",
        "A": "Memungkinkan Indonesia untuk bergantung sepenuhnya kepada bantuan negara lain",
        "B": "Membuka peluang ekspor, transfer teknologi, investasi, dan memperkuat posisi Indonesia di dunia internasional",
        "C": "Memberikan keleluasaan negara lain untuk ikut campur dalam urusan dalam negeri Indonesia",
        "D": "Mengurangi kemandirian Indonesia dalam mengelola sumber daya alamnya",
        "E": "Memaksa Indonesia mengikuti semua kebijakan yang ditetapkan oleh negara mitra",
        "jawaban": "B",
    },
    {
        "no": 15, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Diplomasi Indonesia dalam Hubungan Antar Negara",
        "cp": "Peserta didik mampu memahami peran diplomasi dalam menjalin hubungan Indonesia dengan negara lain",
        "ipk": "Menjelaskan peran diplomasi dan jalur-jalur diplomasi yang digunakan Indonesia dalam hubungan antar negara",
        "indikator": "Disajikan situasi hubungan antar negara, peserta didik dapat menjelaskan peran dan jalur diplomasi yang digunakan Indonesia",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "Dalam menyelesaikan permasalahan dengan negara lain, Indonesia senantiasa mengedepankan jalur diplomasi daripada konfrontasi. Hal ini sesuai dengan tujuan nasional Indonesia dalam UUD 1945 yaitu ...",
        "A": "Membentuk pemerintahan yang kuat dan berwibawa di mata dunia",
        "B": "Mencapai kemakmuran sebesar-besarnya bagi seluruh rakyat Indonesia",
        "C": "Ikut melaksanakan ketertiban dunia berdasarkan kemerdekaan, perdamaian abadi, dan keadilan sosial",
        "D": "Menjadikan Indonesia sebagai negara terkuat di kawasan Asia Tenggara",
        "E": "Mendapatkan pengakuan dari semua negara di dunia sebagai negara besar",
        "jawaban": "C",
    },
    # ---- MATERI 4: NILAI-NILAI PANCASILA DALAM PEMBANGUNAN NASIONAL (Soal 16-20) ----
    {
        "no": 16, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Pancasila sebagai Dasar Pembangunan Nasional",
        "cp": "Peserta didik mampu memahami kedudukan Pancasila sebagai dasar dan pedoman pembangunan nasional",
        "ipk": "Menjelaskan kedudukan Pancasila sebagai dasar dan pedoman dalam pembangunan nasional Indonesia",
        "indikator": "Diberikan pernyataan tentang pembangunan nasional, peserta didik dapat menjelaskan kedudukan Pancasila sebagai dasar dan pedoman pembangunan nasional",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Pancasila memiliki kedudukan yang sangat strategis dalam pembangunan nasional Indonesia. Kedudukan Pancasila dalam pembangunan nasional adalah sebagai ...",
        "A": "Pedoman teknis dalam setiap proyek pembangunan infrastruktur",
        "B": "Dasar filosofis, panduan arah, dan sumber nilai dalam setiap aspek pembangunan nasional",
        "C": "Anggaran dasar yang mengatur sumber pendanaan pembangunan",
        "D": "Petunjuk operasional bagi para pelaksana pembangunan di lapangan",
        "E": "Standar internasional yang harus dipenuhi dalam pembangunan nasional",
        "jawaban": "B",
    },
    {
        "no": 17, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Nilai Sila Pertama dalam Pembangunan Nasional",
        "cp": "Peserta didik mampu menganalisis implementasi nilai sila pertama Pancasila dalam pembangunan nasional",
        "ipk": "Menganalisis implementasi nilai Ketuhanan Yang Maha Esa dalam kebijakan dan program pembangunan nasional",
        "indikator": "Diberikan contoh kebijakan pembangunan, peserta didik dapat menganalisis yang mencerminkan nilai sila pertama Pancasila",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Nilai sila pertama Pancasila, Ketuhanan Yang Maha Esa, tercermin dalam pembangunan nasional melalui ...",
        "A": "Program pembangunan yang hanya mengutamakan kelompok masyarakat tertentu",
        "B": "Pembangunan gedung-gedung pencakar langit sebagai simbol kemajuan bangsa",
        "C": "Jaminan kebebasan beragama, pembangunan sarana ibadah, dan pengembangan moral bangsa dalam setiap program pembangunan",
        "D": "Mewajibkan semua program pembangunan menggunakan simbol-simbol keagamaan",
        "E": "Pembangunan yang hanya difokuskan pada negara-negara yang mayoritas beragama sama",
        "jawaban": "C",
    },
    {
        "no": 18, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Nilai Kemanusiaan dan Persatuan dalam Pembangunan",
        "cp": "Peserta didik mampu menganalisis implementasi nilai sila kedua dan ketiga Pancasila dalam pembangunan nasional",
        "ipk": "Menganalisis implementasi nilai kemanusiaan dan persatuan dalam pembangunan yang berkeadilan dan merata",
        "indikator": "Diberikan contoh program pembangunan, peserta didik dapat menganalisis yang mencerminkan nilai kemanusiaan dan persatuan sila kedua dan ketiga Pancasila",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Program pembangunan nasional yang mencerminkan nilai sila kedua (Kemanusiaan yang Adil dan Beradab) dan sila ketiga (Persatuan Indonesia) Pancasila adalah ...",
        "A": "Pembangunan infrastruktur yang hanya dipusatkan di kota-kota besar",
        "B": "Program pembangunan yang mengutamakan daerah-daerah yang menguntungkan secara ekonomi",
        "C": "Program pembangunan yang merata di seluruh pelosok Indonesia dan menghormati keberagaman budaya lokal",
        "D": "Pembangunan yang hanya memperhatikan kebutuhan kelompok mayoritas",
        "E": "Program pembangunan yang didanai sepenuhnya oleh investor asing",
        "jawaban": "C",
    },
    {
        "no": 19, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Nilai Demokrasi dalam Pembangunan Nasional",
        "cp": "Peserta didik mampu menganalisis implementasi nilai sila keempat Pancasila dalam proses pembangunan",
        "ipk": "Menganalisis penerapan nilai musyawarah mufakat dalam proses perencanaan dan pelaksanaan pembangunan",
        "indikator": "Disajikan proses perencanaan pembangunan, peserta didik dapat menganalisis penerapan nilai musyawarah mufakat sila keempat Pancasila",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Nilai sila keempat Pancasila, Kerakyatan yang Dipimpin oleh Hikmat Kebijaksanaan dalam Permusyawaratan/Perwakilan, tercermin dalam proses pembangunan melalui ...",
        "A": "Keputusan pembangunan diambil sepenuhnya oleh pemimpin tanpa melibatkan rakyat",
        "B": "Musrenbang (Musyawarah Perencanaan Pembangunan) yang melibatkan masyarakat dalam perencanaan program pembangunan daerah",
        "C": "Pembangunan yang hanya mengikuti keinginan investor tanpa mempertimbangkan aspirasi rakyat",
        "D": "Seluruh keputusan pembangunan diserahkan kepada lembaga asing yang lebih ahli",
        "E": "Pembangunan dilaksanakan berdasarkan perintah pusat tanpa memperhatikan kondisi daerah",
        "jawaban": "B",
    },
    {
        "no": 20, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Nilai Keadilan Sosial dalam Pembangunan Nasional",
        "cp": "Peserta didik mampu menganalisis implementasi nilai sila kelima Pancasila dalam pembangunan yang berkeadilan",
        "ipk": "Menganalisis implementasi nilai Keadilan Sosial bagi Seluruh Rakyat Indonesia dalam kebijakan pembangunan nasional",
        "indikator": "Diberikan contoh kebijakan pembangunan, peserta didik dapat menganalisis yang mencerminkan nilai keadilan sosial sila kelima Pancasila",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "Nilai sila kelima Pancasila, Keadilan Sosial bagi Seluruh Rakyat Indonesia, menjadi landasan penting dalam pembangunan nasional. Program pembangunan yang paling mencerminkan nilai sila kelima Pancasila adalah ...",
        "A": "Pembangunan yang hanya difokuskan pada sektor yang menguntungkan bagi negara",
        "B": "Program subsidi dan bantuan sosial yang hanya diberikan kepada kelompok tertentu yang dekat dengan kekuasaan",
        "C": "Pembangunan infrastruktur di daerah terpencil, program bantuan sosial bagi masyarakat miskin, dan pemerataan akses pendidikan dan kesehatan",
        "D": "Prioritas pembangunan pada industri besar yang dapat meningkatkan PDB secara signifikan",
        "E": "Pengalihan seluruh sumber daya pembangunan ke pusat untuk dikelola secara terpusat",
        "jawaban": "C",
    },
]

essay_data = [
    {
        "no": 1, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Hak dan Kewajiban Warga Sekolah",
        "cp": "Peserta didik mampu memahami dan menganalisis hak dan kewajiban sebagai warga sekolah secara komprehensif",
        "ipk": "Menganalisis hak dan kewajiban peserta didik, guru, dan kepala sekolah dalam ekosistem sekolah",
        "indikator": "Diberikan pertanyaan tentang hak dan kewajiban di sekolah, peserta didik dapat menganalisis secara komprehensif hak dan kewajiban semua pihak di lingkungan sekolah",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Sekolah merupakan komunitas belajar yang di dalamnya terdapat berbagai pihak, yaitu peserta didik, guru, dan tenaga kependidikan. Setiap pihak memiliki hak dan kewajiban masing-masing yang saling berkaitan dan mendukung.\n\nKeseimbangan antara hak dan kewajiban di lingkungan sekolah sangat penting untuk menciptakan suasana belajar yang kondusif, nyaman, dan menyenangkan. Ketika setiap warga sekolah melaksanakan kewajibannya dengan baik, maka hak setiap pihak pun akan terpenuhi secara optimal.\n\nSebaliknya, ketika salah satu pihak tidak melaksanakan kewajibannya, hal tersebut dapat mengganggu hak pihak lain dan merusak ekosistem sekolah yang seharusnya harmonis dan produktif.",
        "pertanyaan": "Jelaskan minimal 4 hak dan 4 kewajiban peserta didik di sekolah! Mengapa penting bagi peserta didik untuk memahami dan melaksanakan hak serta kewajibannya di sekolah?",
        "jawaban": "Hak Peserta Didik di Sekolah:\n1. Mendapatkan pengajaran yang berkualitas dari guru yang kompeten\n2. Menggunakan fasilitas sekolah (perpustakaan, laboratorium, lapangan) untuk keperluan belajar\n3. Mendapatkan penilaian yang adil dan objektif sesuai kemampuan\n4. Mendapatkan perlindungan dari tindakan kekerasan, diskriminasi, dan perundungan (bullying)\n5. Menyampaikan pendapat dan aspirasi melalui forum yang tersedia (OSIS, musyawarah kelas)\n\nKewajiban Peserta Didik di Sekolah:\n1. Mengikuti kegiatan pembelajaran dengan penuh semangat dan kedisiplinan\n2. Menghormati guru, tenaga kependidikan, dan sesama peserta didik\n3. Mematuhi tata tertib dan peraturan sekolah yang berlaku\n4. Menjaga kebersihan, ketertiban, dan kelestarian fasilitas sekolah\n5. Mengerjakan tugas dan ulangan dengan jujur tanpa mencontek\n\nPentingnya Memahami Hak dan Kewajiban di Sekolah:\nMemahami hak dan kewajiban di sekolah penting karena:\n- Menciptakan lingkungan belajar yang kondusif dan harmonis\n- Melatih peserta didik untuk bertanggung jawab sejak dini\n- Membangun karakter disiplin yang berguna dalam kehidupan bermasyarakat\n- Memastikan proses pendidikan berjalan dengan efektif dan berkualitas\n- Merupakan bekal memahami hak dan kewajiban dalam kehidupan berbangsa dan bernegara",
    },
    {
        "no": 2, "elemen": "Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
        "materi": "Hak dan Kewajiban Warga Masyarakat",
        "cp": "Peserta didik mampu menganalisis hak dan kewajiban warga masyarakat dalam kehidupan bermasyarakat",
        "ipk": "Menganalisis hak dan kewajiban warga masyarakat berdasarkan norma sosial dan hukum yang berlaku",
        "indikator": "Diberikan kasus kehidupan bermasyarakat, peserta didik dapat menganalisis hak dan kewajiban warga masyarakat serta pentingnya keseimbangannya",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Kehidupan bermasyarakat menuntut setiap anggotanya untuk memahami dan melaksanakan hak serta kewajibannya. Dalam masyarakat yang teratur, hak dan kewajiban bukan sekadar tuntutan hukum, tetapi merupakan bagian dari norma sosial yang mengatur kehidupan bersama.\n\nSetiap warga masyarakat berhak mendapatkan layanan publik yang baik, keamanan, kenyamanan lingkungan, dan diperlakukan secara adil. Namun, untuk mendapatkan hak-hak tersebut, setiap warga pun harus melaksanakan kewajibannya terhadap masyarakat dan negara.\n\nKonsep gotong royong yang menjadi tradisi masyarakat Indonesia sejak dahulu merupakan perwujudan nyata dari keseimbangan hak dan kewajiban dalam kehidupan bermasyarakat. Gotong royong mencerminkan kesadaran bahwa setiap orang memiliki tanggung jawab terhadap kebaikan bersama.",
        "pertanyaan": "Jelaskan minimal 4 hak dan 4 kewajiban yang dimiliki oleh warga masyarakat! Berikan contoh nyata ketika ketidakseimbangan antara hak dan kewajiban menimbulkan masalah dalam kehidupan bermasyarakat!",
        "jawaban": "Hak Warga Masyarakat:\n1. Hak mendapatkan keamanan dan perlindungan dari gangguan keamanan di lingkungan tempat tinggal\n2. Hak mendapatkan lingkungan yang bersih, sehat, dan nyaman untuk dihuni\n3. Hak mendapatkan pelayanan publik yang baik dari pemerintah setempat\n4. Hak berpartisipasi dalam pengambilan keputusan yang menyangkut kepentingan bersama (musyawarah warga)\n5. Hak diperlakukan secara adil tanpa diskriminasi dalam kehidupan bermasyarakat\n\nKewajiban Warga Masyarakat:\n1. Menjaga keamanan dan ketertiban lingkungan dengan aktif dalam kegiatan siskamling\n2. Menjaga kebersihan lingkungan dengan tidak membuang sampah sembarangan\n3. Membayar iuran atau retribusi yang telah disepakati bersama (iuran RT/RW, kebersihan)\n4. Menghormati hak-hak tetangga dan tidak mengganggu ketenangan lingkungan\n5. Ikut berpartisipasi dalam kegiatan gotong royong dan musyawarah warga\n\nContoh Ketidakseimbangan yang Menimbulkan Masalah:\nContoh nyata: Seorang warga yang menuntut keamanan lingkungan (hak) namun menolak ikut serta dalam kegiatan siskamling (kewajiban). Akibatnya:\n- Sistem keamanan lingkungan melemah karena kekurangan personel\n- Menimbulkan kecemburuan sosial dan ketegangan antar warga\n- Keamanan lingkungan menjadi tidak terjamin karena tidak semua warga berkontribusi\n- Rasa kebersamaan dan solidaritas antar warga menjadi berkurang\nSolusi: Menyadarkan warga bahwa hak dan kewajiban harus dilaksanakan secara seimbang demi kebaikan bersama.",
    },
    {
        "no": 3, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Sistem Pertahanan Rakyat Semesta dan Peran Warga Negara",
        "cp": "Peserta didik mampu memahami sistem pertahanan rakyat semesta dan peran warga negara di dalamnya",
        "ipk": "Menganalisis konsep Sishankamrata dan bentuk partisipasi warga negara dalam sistem pertahanan Indonesia",
        "indikator": "Diberikan pertanyaan tentang Sishankamrata, peserta didik dapat menganalisis konsep dan peran warga negara dalam sistem pertahanan rakyat semesta",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Indonesia sebagai negara kepulauan yang luas dengan posisi geografis yang strategis memiliki kekhasan dalam sistem pertahanan dan keamanan negaranya. Sejak awal kemerdekaan, Indonesia menyadari bahwa pertahanan negara tidak bisa hanya mengandalkan kekuatan militer saja, melainkan harus melibatkan seluruh potensi bangsa.\n\nKonsep Sistem Pertahanan Rakyat Semesta (Sishankamrata) yang dianut Indonesia menempatkan rakyat sebagai kekuatan utama yang didukung oleh kekuatan militer. Konsep ini berakar dari pengalaman perjuangan kemerdekaan Indonesia di mana rakyat memiliki peran yang sangat besar dalam mengusir penjajah.\n\nDalam UUD NRI 1945 Pasal 30 ditegaskan bahwa pertahanan dan keamanan negara adalah tanggung jawab seluruh warga negara. Ini berarti setiap warga negara memiliki hak dan kewajiban untuk ikut serta dalam upaya pertahanan dan keamanan negara.",
        "pertanyaan": "Jelaskan konsep Sistem Pertahanan Rakyat Semesta (Sishankamrata)! Uraikan komponen-komponennya dan berikan contoh peran nyata yang dapat dilakukan oleh pelajar SMA dalam mendukung sistem pertahanan dan keamanan negara!",
        "jawaban": "Konsep Sistem Pertahanan Rakyat Semesta (Sishankamrata):\nSishankamrata adalah sistem pertahanan negara Indonesia yang bersifat semesta, artinya melibatkan seluruh warga negara, wilayah, dan sumber daya nasional, serta dipersiapkan secara dini oleh pemerintah dan diselenggarakan secara total, terpadu, terarah, dan berlanjut.\n\nKomponen Sishankamrata:\n1. Komponen Utama: Tentara Nasional Indonesia (TNI) - sebagai kekuatan militer profesional yang bertugas mempertahankan kedaulatan negara\n2. Komponen Cadangan: Warga negara, sumber daya alam, sumber daya buatan, dan sarana prasarana yang dapat dimobilisasi untuk memperbesar dan memperkuat komponen utama\n3. Komponen Pendukung: Seluruh warga negara, sumber daya nasional, dan sarana prasarana yang memberikan dukungan umum terhadap upaya pertahanan\n\nPeran Pelajar SMA dalam Mendukung Sishankamrata:\n1. Mengikuti pelajaran Pendidikan Kewarganegaraan dengan serius untuk memahami nilai-nilai bela negara\n2. Aktif dalam kegiatan Pramuka, PMR, dan ekstrakurikuler yang melatih kedisiplinan dan kepemimpinan\n3. Menjaga kondusivitas lingkungan sekolah dan masyarakat\n4. Melaporkan hal-hal mencurigakan kepada orang yang berwenang\n5. Berprestasi di bidang sains, teknologi, dan olahraga sebagai kontribusi bagi kemajuan bangsa\n6. Menyebarkan semangat cinta tanah air dan bela negara di media sosial",
    },
    {
        "no": 4, "elemen": "Sistem Pertahanan dan Keamanan Negara Indonesia",
        "materi": "Ancaman terhadap Pertahanan dan Keamanan Negara",
        "cp": "Peserta didik mampu menganalisis berbagai bentuk ancaman terhadap pertahanan dan keamanan negara Indonesia",
        "ipk": "Mengidentifikasi dan menganalisis ancaman militer dan non-militer terhadap pertahanan keamanan negara",
        "indikator": "Diberikan berbagai bentuk ancaman, peserta didik dapat mengidentifikasi dan menganalisis ancaman militer dan non-militer terhadap pertahanan keamanan Indonesia",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Pertahanan dan keamanan negara tidak hanya berhadapan dengan ancaman berupa serangan militer dari negara lain. Di era modern ini, ancaman terhadap pertahanan dan keamanan negara telah berkembang menjadi sangat beragam dan kompleks, meliputi ancaman militer dan non-militer.\n\nAncaman non-militer bahkan sering dianggap lebih berbahaya karena sulit dideteksi, berdampak jangka panjang, dan dapat menghancurkan sendi-sendi kehidupan berbangsa dan bernegara tanpa harus menggunakan kekuatan bersenjata. Terorisme, radikalisme, perang siber, dan ancaman ekonomi adalah beberapa contoh ancaman non-militer yang kini semakin nyata.\n\nMenghadapi ragam ancaman ini memerlukan strategi pertahanan yang komprehensif dan melibatkan seluruh komponen bangsa, bukan hanya TNI dan Polri saja.",
        "pertanyaan": "Jelaskan perbedaan antara ancaman militer dan ancaman non-militer terhadap pertahanan dan keamanan negara Indonesia! Berikan masing-masing 3 contoh dan uraikan cara mengatasinya!",
        "jawaban": "Perbedaan Ancaman Militer dan Non-Militer:\n\nAncaman Militer:\nDefinisi: Ancaman yang menggunakan kekuatan bersenjata yang terorganisir yang dinilai mempunyai kemampuan membahayakan kedaulatan negara, keutuhan wilayah, dan keselamatan bangsa.\nCiri-ciri: Menggunakan kekerasan fisik/bersenjata, terorganisir dan terencana, bersumber dari kekuatan militer musuh.\n\nContoh Ancaman Militer:\n1. Agresi atau invasi militer dari negara lain ke wilayah Indonesia\n2. Pemberontakan bersenjata yang ingin memisahkan diri dari NKRI\n3. Sabotase terhadap instalasi vital militer negara\n\nCara Mengatasi: Penguatan TNI, pemberdayaan komponen cadangan dan pendukung, diplomasi pertahanan, dan kerja sama pertahanan dengan negara sahabat.\n\nAncaman Non-Militer:\nDefinisi: Ancaman yang tidak menggunakan kekuatan bersenjata tetapi dapat membahayakan kedaulatan negara, keutuhan wilayah, dan keselamatan bangsa.\nCiri-ciri: Tidak menggunakan kekerasan fisik langsung, dampaknya lebih gradual namun lebih luas dan dalam.\n\nContoh Ancaman Non-Militer:\n1. Terorisme dan radikalisme yang mengancam keamanan dan ideologi bangsa\n2. Serangan siber (cyber attack) terhadap infrastruktur vital negara\n3. Ancaman ekonomi berupa monopoli ekonomi dan praktik ekonomi yang merugikan negara\n\nCara Mengatasi: Deradikalisasi, penguatan keamanan siber nasional (BSSN), kebijakan ekonomi nasional yang kuat, dan pemberdayaan masyarakat dalam mengenali ancaman non-militer.",
    },
    {
        "no": 5, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Politik Luar Negeri Bebas Aktif Indonesia",
        "cp": "Peserta didik mampu menganalisis prinsip dan implementasi politik luar negeri bebas aktif Indonesia",
        "ipk": "Menganalisis prinsip bebas aktif dalam politik luar negeri Indonesia dan implementasinya dalam hubungan antar negara",
        "indikator": "Diberikan informasi tentang politik luar negeri Indonesia, peserta didik dapat menganalisis prinsip bebas aktif dan implementasinya",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Sejak awal kemerdekaan, Indonesia telah menetapkan politik luar negeri yang bebas dan aktif. Prinsip ini pertama kali diperkenalkan oleh Wakil Presiden pertama Indonesia, Mohammad Hatta, dalam pidatonya di depan KNIP (Komite Nasional Indonesia Pusat) pada tahun 1948.\n\nPolitik luar negeri bebas aktif berarti Indonesia tidak berpihak pada blok kekuatan manapun (bebas), namun tetap aktif berperan dalam pergaulan internasional, terutama dalam menjaga perdamaian dunia (aktif). Prinsip ini didasarkan pada semangat kemerdekaan dan keyakinan bahwa Indonesia harus mampu berdiri di atas kaki sendiri.\n\nDalam perkembangannya, politik luar negeri Indonesia telah menghasilkan berbagai pencapaian penting, mulai dari penyelenggaraan Konferensi Asia-Afrika di Bandung (1955), menjadi anggota pendiri ASEAN (1967), hingga memimpin berbagai forum internasional. Prinsip bebas aktif tetap menjadi kompas yang memandu hubungan Indonesia dengan negara-negara lain.",
        "pertanyaan": "Jelaskan makna prinsip 'bebas aktif' dalam politik luar negeri Indonesia! Berikan minimal 4 contoh nyata implementasi prinsip bebas aktif dalam hubungan Indonesia dengan negara lain atau organisasi internasional!",
        "jawaban": "Makna Prinsip Bebas Aktif:\n- Bebas: Indonesia tidak memihak atau terikat pada blok kekuatan manapun (Blok Barat/Amerika atau Blok Timur/Soviet di era Perang Dingin; maupun persekutuan militer tertentu di era kini). Indonesia menentukan sendiri arah kebijakannya berdasarkan kepentingan nasional.\n- Aktif: Indonesia tidak berdiam diri dalam pergaulan internasional, melainkan aktif berperan serta dalam menjaga perdamaian dunia dan menyelesaikan masalah-masalah internasional sesuai dengan kemampuannya.\n\nDasar Hukum: Pembukaan UUD 1945 alinea keempat (ikut melaksanakan ketertiban dunia) dan UU No. 37 Tahun 1999 tentang Hubungan Luar Negeri.\n\nContoh Implementasi Prinsip Bebas Aktif:\n\n1. Menjadi Pemrakarsa dan Tuan Rumah Konferensi Asia-Afrika (1955)\nIndonesia di bawah kepemimpinan Presiden Soekarno memprakarsai dan menjadi tuan rumah Konferensi Asia-Afrika di Bandung yang menghasilkan Dasasila Bandung sebagai landasan hubungan antar bangsa yang damai.\n\n2. Menjadi Anggota Pendiri ASEAN (1967)\nIndonesia bersama Thailand, Malaysia, Filipina, dan Singapura mendirikan ASEAN sebagai organisasi regional untuk menjaga perdamaian dan stabilitas di kawasan Asia Tenggara tanpa terikat blok manapun.\n\n3. Mengirimkan Pasukan Garuda dalam Misi Perdamaian PBB\nIndonesia secara konsisten mengirimkan pasukan TNI dalam misi perdamaian PBB (UNPKF) di berbagai negara konflik seperti Lebanon, Kongo, dan Sudan sebagai wujud kontribusi aktif dalam menjaga perdamaian dunia.\n\n4. Berperan Aktif dalam Forum G20\nIndonesia aktif berpartisipasi dalam forum G20 sebagai negara berkembang yang menyuarakan kepentingan negara-negara Selatan Global dalam isu ekonomi, perubahan iklim, dan pembangunan berkelanjutan.",
    },
    {
        "no": 6, "elemen": "Peran Indonesia dalam Hubungan Antar Negara",
        "materi": "Peran Indonesia dalam Organisasi Internasional",
        "cp": "Peserta didik mampu menganalisis peran dan kontribusi Indonesia dalam berbagai organisasi internasional",
        "ipk": "Menganalisis kontribusi Indonesia dalam PBB, ASEAN, dan organisasi internasional lainnya",
        "indikator": "Diberikan informasi tentang organisasi internasional, peserta didik dapat menganalisis peran dan kontribusi Indonesia di dalamnya",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Indonesia merupakan anggota aktif dari berbagai organisasi internasional, baik di tingkat regional maupun global. Keanggotaan Indonesia dalam organisasi-organisasi ini tidak hanya memberikan manfaat bagi Indonesia sendiri, tetapi juga merupakan wujud tanggung jawab Indonesia sebagai bagian dari komunitas internasional.\n\nDalam berbagai forum internasional, Indonesia kerap menyuarakan kepentingan negara-negara berkembang, memperjuangkan keadilan ekonomi global, mendorong penyelesaian konflik secara damai, dan berkontribusi dalam penanganan isu-isu global seperti perubahan iklim dan terorisme.\n\nPrestasi dan kontribusi Indonesia dalam organisasi internasional merupakan cerminan dari implementasi politik luar negeri bebas aktif yang selama ini dianut. Setiap kontribusi Indonesia di kancah internasional pada akhirnya juga harus memberikan dampak positif bagi kepentingan nasional Indonesia.",
        "pertanyaan": "Uraikan peran dan kontribusi Indonesia dalam minimal 3 organisasi internasional yang berbeda! Jelaskan manfaat yang diperoleh Indonesia dari keanggotaannya dalam organisasi-organisasi internasional tersebut!",
        "jawaban": "Peran Indonesia dalam Organisasi Internasional:\n\n1. Perserikatan Bangsa-Bangsa (PBB)\nPeran Indonesia:\n- Mengirimkan pasukan Garuda dalam misi perdamaian PBB di berbagai negara konflik\n- Aktif dalam sidang-sidang PBB dan berbagai komite PBB\n- Pernah menjabat sebagai anggota tidak tetap Dewan Keamanan PBB\n- Aktif dalam badan-badan PBB seperti WHO, UNESCO, UNICEF, dll.\nManfaat: Indonesia mendapatkan pengakuan internasional, akses terhadap program-program pembangunan PBB, dan forum untuk menyuarakan kepentingan nasional.\n\n2. ASEAN (Association of Southeast Asian Nations)\nPeran Indonesia:\n- Salah satu negara pendiri ASEAN tahun 1967\n- Kerap menjadi penengah dalam penyelesaian konflik di kawasan (misal: konflik Kamboja, Laut China Selatan)\n- Mengajukan berbagai inisiatif untuk penguatan kerja sama ASEAN\n- Menjadi tuan rumah berbagai pertemuan ASEAN\nManfaat: Kemudahan perdagangan, investasi, dan mobilitas warga dalam kawasan ASEAN; keamanan regional yang lebih terjamin.\n\n3. G20 (Group of Twenty)\nPeran Indonesia:\n- Aktif berpartisipasi dalam pertemuan G20 mewakili kepentingan negara berkembang\n- Pada tahun 2022, Indonesia menjadi Presiden G20 dengan tema 'Recover Together, Recover Stronger'\n- Menyuarakan kepentingan ekonomi negara-negara berkembang dalam forum tersebut\nManfaat: Peningkatan reputasi Indonesia di mata dunia, akses terhadap forum pengambilan keputusan ekonomi global, dan peluang investasi dari negara-negara G20.",
    },
    {
        "no": 7, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Implementasi Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "cp": "Peserta didik mampu menganalisis implementasi nilai-nilai Pancasila dalam berbagai aspek pembangunan nasional",
        "ipk": "Menganalisis implementasi nilai-nilai kelima sila Pancasila dalam program-program pembangunan nasional Indonesia",
        "indikator": "Diberikan berbagai program pembangunan nasional, peserta didik dapat menganalisis implementasi nilai-nilai kelima sila Pancasila di dalamnya",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Pembangunan nasional Indonesia bukan sekadar pembangunan fisik seperti infrastruktur dan fasilitas umum, tetapi merupakan upaya menyeluruh untuk meningkatkan kualitas kehidupan manusia Indonesia dalam segala aspek. Pembangunan ini harus dilaksanakan berdasarkan nilai-nilai Pancasila sebagai dasar negara dan pandangan hidup bangsa.\n\nNilai-nilai kelima sila Pancasila memberikan panduan yang komprehensif tentang bagaimana pembangunan harus direncanakan, dilaksanakan, dan dievaluasi. Mulai dari aspek religiositas, kemanusiaan, persatuan, demokrasi, hingga keadilan sosial, semua tercakup dalam nilai-nilai Pancasila.\n\nDengan menjadikan Pancasila sebagai landasan pembangunan, Indonesia berupaya menciptakan masyarakat yang tidak hanya maju secara ekonomi dan teknologi, tetapi juga bermartabat, berkeadilan, dan memiliki jati diri yang kuat sebagai bangsa Indonesia.",
        "pertanyaan": "Jelaskan bagaimana nilai-nilai kelima sila Pancasila dapat diimplementasikan dalam pembangunan nasional Indonesia! Berikan contoh program pembangunan konkret yang mencerminkan nilai setiap sila!",
        "jawaban": "Implementasi Nilai-Nilai Pancasila dalam Pembangunan Nasional:\n\n1. Sila Pertama: Ketuhanan Yang Maha Esa\nImplementasi: Pembangunan harus menjamin kebebasan beragama dan mendukung kehidupan spiritual masyarakat.\nContoh Program: Pembangunan sarana ibadah di seluruh pelosok Indonesia; program Kementerian Agama dalam bimbingan keagamaan; integrasi pendidikan agama dalam kurikulum nasional; kebijakan yang mengakui dan menghormati semua agama yang diakui negara.\n\n2. Sila Kedua: Kemanusiaan yang Adil dan Beradab\nImplementasi: Pembangunan harus menghargai harkat dan martabat manusia serta memberikan perlakuan yang adil bagi semua warga.\nContoh Program: Program jaminan kesehatan nasional (BPJS Kesehatan); program bantuan sosial bagi kelompok rentan; pembangunan rumah susun bagi masyarakat berpenghasilan rendah; penegakan hukum HAM.\n\n3. Sila Ketiga: Persatuan Indonesia\nImplementasi: Pembangunan harus memperkuat persatuan dan kesatuan bangsa, menghargai keberagaman, dan tidak menimbulkan kesenjangan antar daerah.\nContoh Program: Pembangunan infrastruktur konektivitas (tol laut, Trans Papua, Trans Kalimantan); program transmigrasi untuk pemerataan penduduk; Festival Budaya Nusantara; program satu harga BBM untuk seluruh Indonesia.\n\n4. Sila Keempat: Kerakyatan yang Dipimpin oleh Hikmat Kebijaksanaan dalam Permusyawaratan/Perwakilan\nImplementasi: Pembangunan harus melibatkan partisipasi rakyat dan melalui proses musyawarah yang demokratis.\nContoh Program: Musrenbang (Musyawarah Perencanaan Pembangunan) di semua tingkatan; Dana Desa yang dikelola secara demokratis; mekanisme pengaduan publik terhadap program pembangunan; keterbukaan informasi publik.\n\n5. Sila Kelima: Keadilan Sosial bagi Seluruh Rakyat Indonesia\nImplementasi: Pembangunan harus menghasilkan distribusi kesejahteraan yang merata dan mengurangi ketimpangan.\nContoh Program: Program Keluarga Harapan (PKH); KIP (Kartu Indonesia Pintar) untuk akses pendidikan; program bedah rumah bagi masyarakat miskin; pembangunan infrastruktur di daerah 3T (Tertinggal, Terdepan, Terluar).",
    },
    {
        "no": 8, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Tantangan Implementasi Nilai Pancasila dalam Pembangunan",
        "cp": "Peserta didik mampu menganalisis tantangan dalam mengimplementasikan nilai-nilai Pancasila dalam pembangunan nasional",
        "ipk": "Menganalisis tantangan dan hambatan dalam mengimplementasikan nilai-nilai Pancasila dalam pembangunan nasional",
        "indikator": "Diberikan situasi pembangunan Indonesia, peserta didik dapat menganalisis tantangan implementasi nilai Pancasila dan solusi untuk mengatasinya",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Mewujudkan pembangunan nasional yang berlandaskan nilai-nilai Pancasila bukanlah tugas yang mudah. Indonesia menghadapi berbagai tantangan kompleks yang dapat menghambat terwujudnya cita-cita pembangunan yang adil, merata, dan bermartabat.\n\nKorupsi, kesenjangan ekonomi, degradasi moral, dan polarisasi sosial merupakan tantangan nyata yang harus dihadapi dalam pembangunan. Tantangan-tantangan ini dapat menggerus nilai-nilai Pancasila jika tidak ditangani dengan serius dan komprehensif.\n\nDi sisi lain, tantangan era globalisasi dan kemajuan teknologi juga membawa dampak ganda bagi pembangunan. Di satu sisi, teknologi membuka peluang percepatan pembangunan, namun di sisi lain juga membawa ancaman berupa pergeseran nilai dan budaya yang dapat melemahkan jati diri bangsa.\n\nMenghadapi semua tantangan ini, komitmen untuk terus menjaga Pancasila sebagai pedoman pembangunan menjadi semakin penting dan mendesak.",
        "pertanyaan": "Identifikasi dan jelaskan minimal 4 tantangan utama yang dihadapi dalam mengimplementasikan nilai-nilai Pancasila dalam pembangunan nasional Indonesia! Untuk setiap tantangan, uraikan solusi yang dapat dilakukan!",
        "jawaban": "Tantangan Implementasi Nilai Pancasila dalam Pembangunan Nasional:\n\n1. Korupsi dan Penyimpangan dalam Pelaksanaan Pembangunan\nTantangan: Korupsi yang masih merajalela dalam pengelolaan anggaran pembangunan bertentangan dengan sila kedua (kemanusiaan) dan kelima (keadilan sosial) Pancasila, karena merugikan rakyat dan menghambat pembangunan yang berkeadilan.\nSolusi: Penguatan lembaga pemberantasan korupsi (KPK), sistem pengawasan berbasis teknologi, transparansi penggunaan anggaran publik, dan pendidikan antikorupsi sejak dini.\n\n2. Kesenjangan Ekonomi dan Ketimpangan Pembangunan\nTantangan: Kesenjangan yang masih lebar antara kelompok kaya dan miskin, serta antara daerah maju dan tertinggal, bertentangan dengan nilai keadilan sosial sila kelima Pancasila.\nSolusi: Kebijakan afirmatif untuk daerah 3T, program redistribusi pendapatan (bantuan sosial, subsidi tepat sasaran), pengembangan infrastruktur di daerah terpencil, dan pemberdayaan UMKM lokal.\n\n3. Degradasi Moral dan Lunturnya Nilai-Nilai Pancasila\nTantangan: Globalisasi dan kemajuan teknologi membawa nilai-nilai asing yang dapat melemahkan karakter bangsa yang berlandaskan Pancasila, terutama di kalangan generasi muda.\nSolusi: Penguatan pendidikan karakter Pancasila di sekolah, revitalisasi budaya lokal, program deradikalisasi, dan pemanfaatan media sosial untuk mempromosikan nilai-nilai Pancasila.\n\n4. Polarisasi Sosial dan Konflik Horizontal\nTantangan: Perpecahan masyarakat berdasarkan garis SARA (Suku, Agama, Ras, dan Antargolongan) mengancam nilai persatuan sila ketiga dan kemanusiaan sila kedua Pancasila.\nSolusi: Penguatan dialog antar kelompok, penegakan hukum yang tegas terhadap ujaran kebencian dan provokasi, program integrasi sosial, serta pembangunan yang berkeadilan untuk mengurangi kecemburuan antar kelompok.",
    },
    {
        "no": 9, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Peran Generasi Muda dalam Pembangunan Nasional Berbasis Pancasila",
        "cp": "Peserta didik mampu merumuskan peran generasi muda dalam mendukung pembangunan nasional yang berlandaskan nilai-nilai Pancasila",
        "ipk": "Merancang kontribusi generasi muda dalam pembangunan nasional yang berlandaskan nilai-nilai Pancasila",
        "indikator": "Diberikan tantangan pembangunan nasional, peserta didik dapat merumuskan peran generasi muda dalam mendukung pembangunan berbasis Pancasila",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Generasi muda Indonesia adalah penentu masa depan bangsa. Sebagai penerus estafet kepemimpinan, generasi muda memiliki tanggung jawab besar untuk memastikan bahwa pembangunan nasional Indonesia terus berjalan di atas rel nilai-nilai Pancasila.\n\nDi era revolusi industri 4.0 dan society 5.0 ini, generasi muda Indonesia berhadapan dengan tantangan sekaligus peluang yang belum pernah ada sebelumnya. Kemampuan teknologi, kreativitas, dan semangat yang dimiliki generasi muda merupakan modal berharga yang harus dioptimalkan untuk pembangunan bangsa.\n\nNamun, semua potensi tersebut hanya akan bermakna jika dilandasi oleh karakter yang kuat berdasarkan nilai-nilai Pancasila. Generasi muda yang berkarakter Pancasila adalah generasi yang tidak hanya cerdas dan terampil, tetapi juga jujur, bertanggung jawab, peduli sesama, dan mencintai tanah air.\n\nOleh karena itu, pembentukan karakter generasi muda yang sesuai dengan nilai-nilai Pancasila merupakan investasi terpenting dalam pembangunan nasional Indonesia.",
        "pertanyaan": "Jelaskan mengapa peran generasi muda sangat penting dalam pembangunan nasional yang berlandaskan Pancasila! Uraikan minimal 5 kontribusi konkret yang dapat diberikan oleh pelajar SMA dalam mendukung pembangunan nasional berbasis nilai-nilai Pancasila!",
        "jawaban": "Pentingnya Peran Generasi Muda dalam Pembangunan Nasional Berbasis Pancasila:\n\nGenerasi muda adalah agen perubahan (agent of change) dan penerus kepemimpinan bangsa. Mereka yang akan mewarisi Indonesia di masa mendatang sehingga kualitas, karakter, dan semangat mereka menentukan kualitas Indonesia di masa depan. Generasi muda yang berkarakter Pancasila akan menjamin kesinambungan pembangunan yang berlandaskan nilai-nilai luhur bangsa.\n\nKontribusi Konkret Pelajar SMA dalam Pembangunan Nasional Berbasis Pancasila:\n\n1. Berprestasi Akademik dan Non-Akademik\nMenjadi yang terbaik di bidang sains, teknologi, seni, olahraga, dan bidang lainnya sebagai investasi sumber daya manusia berkualitas yang dibutuhkan dalam pembangunan nasional (mencerminkan sila kedua dan ketiga Pancasila).\n\n2. Membangun Jiwa Wirausaha Berbasis Nilai Pancasila\nMengembangkan usaha kecil berbasis produk lokal, membantu memasarkan produk UMKM masyarakat sekitar melalui platform digital, menciptakan lapangan kerja yang berkontribusi pada keadilan sosial (mencerminkan sila kelima).\n\n3. Aktif dalam Kegiatan Sosial dan Kemanusiaan\nBerpartisipasi dalam kegiatan bakti sosial, penggalangan dana untuk korban bencana, program literasi untuk anak-anak kurang mampu, sebagai wujud nilai kemanusiaan sila kedua Pancasila.\n\n4. Menjadi Agen Anti-Korupsi di Lingkungan Sekolah\nMenerapkan nilai kejujuran dalam ulangan dan tugas, aktif dalam program sekolah bebas korupsi, mengajak teman-teman untuk menolak praktik suap dan gratifikasi (mencerminkan nilai kemanusiaan dan keadilan sosial).\n\n5. Memanfaatkan Teknologi Digital untuk Kebaikan Bangsa\nMenyebarkan konten positif tentang keberagaman budaya Indonesia, melawan hoaks dan ujaran kebencian, mempromosikan produk dalam negeri, dan menggunakan media sosial untuk kampanye nilai-nilai Pancasila kepada sesama generasi muda.",
    },
    {
        "no": 10, "elemen": "Nilai-Nilai Pancasila dalam Pembangunan Nasional",
        "materi": "Pancasila sebagai Pedoman Pembangunan Berkeadilan",
        "cp": "Peserta didik mampu menganalisis Pancasila sebagai pedoman dalam mewujudkan pembangunan yang berkeadilan bagi seluruh rakyat",
        "ipk": "Menganalisis relevansi nilai-nilai Pancasila sebagai pedoman dalam mewujudkan pembangunan nasional yang berkeadilan",
        "indikator": "Diberikan kondisi kesenjangan pembangunan, peserta didik dapat menganalisis relevansi Pancasila sebagai pedoman pembangunan berkeadilan",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Salah satu tantangan terbesar pembangunan Indonesia adalah mewujudkan keadilan sosial bagi seluruh rakyat Indonesia, sebagaimana diamanatkan oleh sila kelima Pancasila. Meskipun Indonesia telah mengalami pertumbuhan ekonomi yang signifikan, masih terdapat kesenjangan yang cukup lebar antara kelompok kaya dan miskin, serta antara daerah maju dan daerah tertinggal.\n\nNilai-nilai Pancasila, jika benar-benar dijadikan pedoman dalam perumusan dan pelaksanaan kebijakan pembangunan, dapat menjadi solusi atas berbagai permasalahan ketimpangan ini. Pancasila menekankan bahwa pembangunan bukan hanya tentang pertumbuhan ekonomi semata, tetapi tentang peningkatan kualitas manusia secara holistik yang mencakup aspek spiritual, sosial, budaya, dan ekonomi.\n\nKonsep 'Tujuan Pembangunan Berkelanjutan' (Sustainable Development Goals/SDGs) yang saat ini diadopsi oleh dunia internasional sesungguhnya selaras dengan nilai-nilai Pancasila yang telah lebih dahulu menekankan pentingnya pembangunan yang adil, berkelanjutan, dan berpusat pada manusia.",
        "pertanyaan": "Jelaskan bagaimana nilai-nilai Pancasila dapat menjadi pedoman dalam mewujudkan pembangunan nasional yang berkeadilan bagi seluruh rakyat Indonesia! Uraikan juga mengapa pendekatan pembangunan berbasis Pancasila dianggap lebih tepat dibandingkan pendekatan pembangunan yang semata-mata berorientasi pada pertumbuhan ekonomi!",
        "jawaban": "Pancasila sebagai Pedoman Pembangunan Berkeadilan:\n\nNilai-nilai Pancasila memberikan kerangka yang komprehensif dan holistik dalam pembangunan nasional:\n\n1. Sila Pertama mengingatkan bahwa pembangunan harus memperhatikan dimensi spiritual manusia, bukan hanya materi. Pembangunan yang menghilangkan nilai religiusitas tidak sesuai dengan karakter bangsa Indonesia.\n\n2. Sila Kedua menegaskan bahwa setiap kebijakan pembangunan harus menghargai martabat manusia dan memberikan perlakuan yang adil tanpa diskriminasi. Pembangunan yang mengeksploitasi manusia tidak sesuai dengan nilai kemanusiaan.\n\n3. Sila Ketiga mengingatkan bahwa pembangunan harus memperkuat persatuan dan tidak boleh menimbulkan perpecahan. Pembangunan yang hanya menguntungkan kelompok tertentu dapat merusak persatuan.\n\n4. Sila Keempat menekankan bahwa proses pembangunan harus demokratis dan partisipatif. Kebijakan pembangunan yang ditentukan secara sepihak tanpa melibatkan rakyat tidak sesuai dengan nilai demokrasi.\n\n5. Sila Kelima mengharuskan bahwa hasil pembangunan harus dinikmati secara adil oleh seluruh rakyat, bukan hanya kelompok tertentu.\n\nKeunggulan Pendekatan Pembangunan Berbasis Pancasila:\nPendekatan pembangunan berbasis Pancasila lebih tepat dibandingkan pembangunan yang semata-mata berorientasi pada pertumbuhan ekonomi karena:\n\n1. Lebih Holistik: Tidak hanya mengukur keberhasilan dari PDB atau pertumbuhan ekonomi, tetapi juga dari kualitas kehidupan spiritual, sosial, budaya, dan lingkungan masyarakat.\n\n2. Lebih Berkeadilan: Memastikan bahwa manfaat pembangunan tidak hanya dinikmati oleh segelintir orang, tetapi merata bagi seluruh rakyat termasuk yang berada di daerah terpencil.\n\n3. Lebih Berkelanjutan: Dengan memperhatikan nilai persatuan dan kemanusiaan, pembangunan berbasis Pancasila lebih mementingkan keseimbangan sosial-ekologis jangka panjang daripada keuntungan ekonomi jangka pendek.\n\n4. Lebih Sesuai dengan Identitas Bangsa: Pancasila berakar dari nilai-nilai luhur bangsa Indonesia, sehingga pendekatan pembangunan berbasis Pancasila lebih sesuai dengan karakter, budaya, dan kebutuhan masyarakat Indonesia.",
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
    for r in [1, 2]:
        for c in [1, 2, 3, 4]:
            try:
                cell = table.rows[r].cells[c]
                if id(cell._tc) != id(table.rows[0].cells[0]._tc):
                    rebuild_cell(cell, f"Buku Acuan/Referensi:\n{REF}")
            except:
                pass
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
            rebuild_cell(c8, str(q.get('nilai','5')))
    except: pass
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
        (1,"Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
         "Peserta didik mampu memahami, membedakan, dan menerapkan hak dan kewajiban sebagai warga sekolah dan masyarakat",
         "Pengertian hak dan kewajiban, pelaksanaan hak dan kewajiban, keseimbangan hak dan kewajiban",
         "Mengidentifikasi dan menerapkan hak dan kewajiban warga sekolah dan masyarakat",
         "5","1 - 5","Pilihan Ganda","Rendah - Sedang"),
        (2,"Sistem Pertahanan dan Keamanan Negara Indonesia",
         "Peserta didik mampu memahami dan menganalisis sistem pertahanan dan keamanan negara Indonesia",
         "Sishankamrata, komponen pertahanan, peran warga negara, bela negara",
         "Menjelaskan konsep Sishankamrata dan peran warga negara dalam pertahanan keamanan",
         "5","6 - 10","Pilihan Ganda","Rendah - Sedang"),
        (3,"Peran Indonesia dalam Hubungan Antar Negara",
         "Peserta didik mampu menganalisis peran dan kontribusi Indonesia dalam hubungan antar negara",
         "Politik luar negeri bebas aktif, peran dalam ASEAN, PBB, kerja sama internasional",
         "Menganalisis prinsip dan implementasi peran Indonesia dalam hubungan antar negara",
         "5","11 - 15","Pilihan Ganda","Rendah - Sulit"),
        (4,"Nilai-Nilai Pancasila dalam Pembangunan Nasional",
         "Peserta didik mampu menganalisis implementasi nilai-nilai Pancasila dalam pembangunan nasional Indonesia",
         "Pancasila sebagai dasar pembangunan, implementasi nilai per sila, pembangunan berkeadilan",
         "Menganalisis implementasi nilai-nilai Pancasila dalam berbagai aspek pembangunan nasional",
         "5","16 - 20","Pilihan Ganda","Rendah - Sulit"),
    ]
    essay_groups = [
        (1,"Hak dan Kewajiban sebagai Warga Sekolah dan Masyarakat",
         "Peserta didik mampu menganalisis hak dan kewajiban warga sekolah dan masyarakat secara mendalam",
         "Hak-kewajiban di sekolah, hak-kewajiban di masyarakat, keseimbangan hak-kewajiban",
         "Menganalisis hak dan kewajiban warga sekolah dan masyarakat beserta dampak ketidakseimbangannya",
         "2","1 - 2","Uraian/Essay","Sedang"),
        (2,"Sistem Pertahanan dan Keamanan Negara Indonesia",
         "Peserta didik mampu menganalisis Sishankamrata dan ancaman terhadap pertahanan keamanan negara",
         "Sishankamrata, ancaman militer dan non-militer, peran warga negara dan pelajar",
         "Menganalisis konsep dan komponen Sishankamrata serta ancaman militer dan non-militer",
         "2","3 - 4","Uraian/Essay","Sedang - Sulit"),
        (3,"Peran Indonesia dalam Hubungan Antar Negara",
         "Peserta didik mampu menganalisis politik luar negeri bebas aktif dan peran Indonesia di organisasi internasional",
         "Prinsip bebas aktif, Konferensi Asia-Afrika, ASEAN, PBB, G20",
         "Menganalisis implementasi bebas aktif dan kontribusi Indonesia dalam organisasi internasional",
         "2","5 - 6","Uraian/Essay","Sedang - Sulit"),
        (4,"Nilai-Nilai Pancasila dalam Pembangunan Nasional",
         "Peserta didik mampu menganalisis implementasi, tantangan, dan relevansi Pancasila dalam pembangunan",
         "Implementasi per sila, tantangan pembangunan, peran generasi muda, pembangunan berkeadilan",
         "Menganalisis implementasi Pancasila dalam pembangunan dan merumuskan peran generasi muda",
         "4","7 - 10","Uraian/Essay","Sedang - Sulit"),
    ]
    for no,elemen,cp,materi,tp,jml,nomor,bentuk,tingkat in pg_groups:
        row = table.rows[no]
        for c,v in enumerate([str(no),elemen,cp,materi,tp,jml,nomor,bentuk,tingkat]):
            try: rebuild_cell(row.cells[c], v)
            except: pass
    for no,elemen,cp,materi,tp,jml,nomor,bentuk,tingkat in essay_groups:
        row = table.rows[no+6]
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

    out = '/home/user/daw/KARTU_SOAL_UAS_GENAP_PKN_X.docx'
    doc.save(out)
    print(f"\nSelesai! Tersimpan: {out}")

main()
