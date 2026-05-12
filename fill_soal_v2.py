from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy

# ============================================================
# KONSTANTA
# ============================================================
NAMA_PENULIS = "Roseno Afandi"
MAPEL = "Komputer Akuntansi"
KELAS = "XI"
REF = "Modul Komputer Akuntansi Accurate Online Kelas XI, Direktorat PSMK Kemdikbud 2023"

# ============================================================
# DATA SOAL PG (20 soal)
# ============================================================
pg_data = [
    {
        "no": 1, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pengenalan Accurate Online",
        "cp": "Peserta didik mampu menjelaskan karakteristik dan keunggulan Accurate Online sebagai software akuntansi berbasis cloud",
        "ipk": "Mengidentifikasi keunggulan Accurate Online berbasis cloud dibandingkan software akuntansi konvensional",
        "indikator": "Disajikan informasi tentang karakteristik Accurate Online, peserta didik dapat menentukan keunggulan utamanya",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Accurate Online adalah aplikasi akuntansi berbasis cloud yang dikembangkan oleh PT Cipta Piranti Sejahtera. Accurate Online dapat diakses melalui berbagai perangkat tanpa perlu instalasi di setiap komputer pengguna. Berdasarkan karakteristik tersebut, salah satu keunggulan utama Accurate Online dibandingkan software akuntansi konvensional adalah ...",
        "A": "Memerlukan instalasi di setiap komputer yang digunakan",
        "B": "Data tersimpan secara lokal sehingga lebih aman dari serangan internet",
        "C": "Dapat diakses kapan saja dan di mana saja menggunakan koneksi internet",
        "D": "Hanya bisa digunakan oleh satu pengguna dalam satu waktu",
        "E": "Tidak memerlukan koneksi internet sama sekali",
        "jawaban": "C",
    },
    {
        "no": 2, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Setup Data Awal Perusahaan",
        "cp": "Peserta didik mampu melakukan setup data awal perusahaan di Accurate Online secara benar",
        "ipk": "Melakukan prosedur pembuatan database perusahaan baru di Accurate Online",
        "indikator": "Diberikan skenario pembuatan perusahaan baru, peserta didik dapat mengidentifikasi informasi yang diperlukan dalam setup awal",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Saat pertama kali menggunakan Accurate Online untuk membuat database perusahaan baru, pengguna harus mengisi berbagai informasi dasar. Informasi ini akan menjadi landasan dari semua transaksi yang akan dicatat. Berikut ini yang BUKAN merupakan informasi yang diisi pada tahap setup awal perusahaan di Accurate Online adalah ...",
        "A": "Nama perusahaan dan alamat perusahaan",
        "B": "Periode fiskal (tahun buku) perusahaan",
        "C": "Metode pencatatan persediaan (FIFO/Average)",
        "D": "Nama dan biodata lengkap direktur utama perusahaan",
        "E": "Mata uang fungsional yang digunakan",
        "jawaban": "D",
    },
    {
        "no": 3, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Daftar Akun (Chart of Accounts)",
        "cp": "Peserta didik mampu membuat dan mengelola daftar akun di Accurate Online",
        "ipk": "Mengidentifikasi pengelompokan akun dalam Accurate Online berdasarkan kode akun",
        "indikator": "Diberikan kode akun tertentu, peserta didik dapat menentukan kelompok akun yang dimaksud",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Dalam Accurate Online, daftar akun yang digunakan untuk mencatat transaksi keuangan perusahaan dikelompokkan berdasarkan kode angka yang sistematis. Setiap kelompok akun memiliki kode awal yang berbeda. Jika kode akun dimulai dengan angka 1, maka kelompok akun tersebut termasuk dalam kelompok ...",
        "A": "Liabilitas (Kewajiban)",
        "B": "Ekuitas (Modal)",
        "C": "Pendapatan",
        "D": "Beban",
        "E": "Aset (Harta)",
        "jawaban": "E",
    },
    {
        "no": 4, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Saldo Awal Akun",
        "cp": "Peserta didik mampu menginput saldo awal akun di Accurate Online untuk perusahaan yang sudah berjalan",
        "ipk": "Menginput saldo awal seluruh akun neraca di Accurate Online",
        "indikator": "Diberikan kasus perusahaan yang akan mulai menggunakan Accurate Online, peserta didik dapat menentukan menu yang tepat untuk input saldo awal",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Sebuah perusahaan dagang yang telah berjalan selama 2 tahun akan mulai menggunakan Accurate Online. Sebelum mencatat transaksi baru, perusahaan harus memasukkan saldo awal untuk seluruh akun yang dimiliki. Di Accurate Online, menu yang digunakan untuk menginput saldo awal akun-akun neraca adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Buku Besar → Saldo Awal",
        "C": "Menu Pembelian → Faktur Pembelian",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Laporan → Neraca",
        "jawaban": "B",
    },
    {
        "no": 5, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Data Master Pelanggan",
        "cp": "Peserta didik mampu membuat dan mengelola data master pelanggan di Accurate Online",
        "ipk": "Membuat data master pelanggan (customer) di Accurate Online sebelum mencatat transaksi penjualan",
        "indikator": "Diberikan skenario transaksi dengan pelanggan baru, peserta didik dapat mengidentifikasi menu yang tepat untuk input data pelanggan",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Sebelum mencatat transaksi penjualan kepada pelanggan baru, seorang akuntan harus terlebih dahulu membuat data master pelanggan tersebut. Tanpa data master yang lengkap, transaksi tidak dapat dicatat dengan benar. Di Accurate Online, data master pelanggan (customer) disimpan dan dikelola melalui menu ...",
        "A": "Menu Pembelian → Data Pemasok",
        "B": "Menu Persediaan → Data Barang",
        "C": "Menu Penjualan → Data Pelanggan",
        "D": "Menu Buku Besar → Daftar Akun",
        "E": "Menu Laporan → Laporan Penjualan",
        "jawaban": "C",
    },
    {
        "no": 6, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Transaksi Penjualan Kredit",
        "cp": "Peserta didik mampu mencatat transaksi penjualan kredit di Accurate Online dengan benar",
        "ipk": "Mencatat faktur penjualan kredit di Accurate Online",
        "indikator": "Diberikan data transaksi penjualan kredit, peserta didik dapat menentukan menu yang tepat untuk mencatatnya",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "PT Maju Jaya melakukan penjualan barang dagangan kepada CV Berkah senilai Rp10.000.000 secara kredit dengan syarat pembayaran Net 30. Akuntan PT Maju Jaya harus mencatat transaksi ini di Accurate Online. Menu yang tepat digunakan untuk mencatat transaksi penjualan kredit tersebut adalah ...",
        "A": "Menu Pembelian → Purchase Order",
        "B": "Menu Penjualan → Faktur Penjualan",
        "C": "Menu Kas & Bank → Penerimaan Kas",
        "D": "Menu Buku Besar → Jurnal Umum",
        "E": "Menu Persediaan → Transfer Stok",
        "jawaban": "B",
    },
    {
        "no": 7, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Data Master Barang / Persediaan",
        "cp": "Peserta didik mampu membuat data master barang dan mengelola persediaan di Accurate Online",
        "ipk": "Membuat data master barang (item list) di Accurate Online",
        "indikator": "Diberikan kondisi sebelum transaksi penjualan/pembelian, peserta didik dapat menentukan menu input data barang",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "Sebelum mencatat transaksi penjualan atau pembelian barang, akuntan harus memastikan data master barang tersebut sudah tersedia di sistem. Tanpa data master barang, sistem tidak dapat memproses transaksi dengan baik. Di Accurate Online, untuk membuat data master barang (item list) dilakukan melalui ...",
        "A": "Menu Penjualan → Data Pelanggan",
        "B": "Menu Pembelian → Data Pemasok",
        "C": "Menu Persediaan → Data Barang",
        "D": "Menu Buku Besar → Daftar Akun",
        "E": "Menu Laporan → Laporan Persediaan",
        "jawaban": "C",
    },
    {
        "no": 8, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Transaksi Pembelian Kredit",
        "cp": "Peserta didik mampu mencatat transaksi pembelian kredit di Accurate Online dengan benar",
        "ipk": "Mencatat faktur pembelian kredit dari pemasok di Accurate Online",
        "indikator": "Diberikan data transaksi pembelian kredit, peserta didik dapat menentukan menu yang tepat untuk mencatatnya",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "CV Berkah Sejahtera melakukan pembelian bahan baku dari PT Supplier Jaya secara kredit senilai Rp15.000.000 dengan syarat pembayaran 2/10 n/30. Untuk mencatat transaksi pembelian kredit ini dengan benar di Accurate Online, menu yang harus digunakan adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Kas & Bank → Pengeluaran Kas",
        "C": "Menu Pembelian → Faktur Pembelian",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Buku Besar → Jurnal Umum",
        "jawaban": "C",
    },
    {
        "no": 9, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Penjualan",
        "cp": "Peserta didik mampu mencatat transaksi retur penjualan di Accurate Online",
        "ipk": "Mencatat dokumen retur penjualan di Accurate Online sesuai referensi faktur aslinya",
        "indikator": "Diberikan skenario pengembalian barang dari pelanggan, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Pelanggan CV Makmur Abadi mengembalikan 5 unit barang yang telah dibeli dari PT Jaya Raya karena barang tersebut tidak sesuai dengan spesifikasi yang dipesan. Untuk mencatat transaksi pengembalian barang dari pelanggan ini di Accurate Online, menu yang digunakan adalah ...",
        "A": "Menu Pembelian → Retur Pembelian",
        "B": "Menu Penjualan → Retur Penjualan",
        "C": "Menu Persediaan → Penyesuaian Stok",
        "D": "Menu Kas & Bank → Pengeluaran Kas",
        "E": "Menu Buku Besar → Jurnal Penyesuaian",
        "jawaban": "B",
    },
    {
        "no": 10, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Pembelian",
        "cp": "Peserta didik mampu mencatat transaksi retur pembelian di Accurate Online",
        "ipk": "Mencatat dokumen retur pembelian kepada pemasok di Accurate Online",
        "indikator": "Diberikan skenario pengembalian barang kepada pemasok, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "PT Sumber Makmur menerima kiriman barang dari pemasoknya, namun setelah diperiksa sebanyak 10 unit barang ditemukan dalam kondisi rusak dan tidak layak digunakan. PT Sumber Makmur memutuskan untuk mengembalikan barang tersebut kepada pemasok. Di Accurate Online, transaksi retur pembelian ini dicatat melalui ...",
        "A": "Menu Penjualan → Retur Penjualan",
        "B": "Menu Pembelian → Retur Pembelian",
        "C": "Menu Persediaan → Transfer Stok",
        "D": "Menu Kas & Bank → Penerimaan Kas",
        "E": "Menu Buku Besar → Jurnal Umum",
        "jawaban": "B",
    },
    {
        "no": 11, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Saldo Awal Persediaan Barang",
        "cp": "Peserta didik mampu menginput saldo awal persediaan barang dagangan di Accurate Online",
        "ipk": "Menginput saldo awal stok barang menggunakan menu penyesuaian stok di Accurate Online",
        "indikator": "Diberikan kasus perusahaan yang baru menggunakan Accurate Online dan memiliki stok barang, peserta didik dapat menentukan menu input saldo awal stok",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Sebuah perusahaan dagang yang baru pertama kali menggunakan Accurate Online memiliki stok barang senilai Rp50.000.000 yang perlu dimasukkan ke dalam sistem. Input saldo awal stok barang dagangan ini di Accurate Online dilakukan melalui menu ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Pembelian → Faktur Pembelian",
        "C": "Menu Persediaan → Penyesuaian Stok (Stock Adjustment)",
        "D": "Menu Buku Besar → Saldo Awal",
        "E": "Menu Laporan → Laporan Persediaan",
        "jawaban": "C",
    },
    {
        "no": 12, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Penerimaan Pembayaran Piutang",
        "cp": "Peserta didik mampu mencatat penerimaan pembayaran dari pelanggan di Accurate Online",
        "ipk": "Mencatat penerimaan pembayaran piutang dari pelanggan menggunakan menu Customer Receipt",
        "indikator": "Diberikan data pembayaran piutang dari pelanggan, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Pelanggan PT Harapan Jaya melakukan pembayaran atas piutangnya kepada PT Maju Bersama sebesar Rp8.000.000 melalui transfer bank. Akuntan PT Maju Bersama harus mencatat penerimaan pembayaran ini di Accurate Online. Pencatatan penerimaan pembayaran piutang dari pelanggan dilakukan melalui menu ...",
        "A": "Menu Pembelian → Pembayaran Pemasok",
        "B": "Menu Penjualan → Penerimaan Pelanggan (Customer Receipt)",
        "C": "Menu Persediaan → Transfer Stok",
        "D": "Menu Kas & Bank → Transfer Kas",
        "E": "Menu Buku Besar → Jurnal Umum",
        "jawaban": "B",
    },
    {
        "no": 13, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pembayaran Utang kepada Pemasok",
        "cp": "Peserta didik mampu mencatat pembayaran utang kepada pemasok di Accurate Online",
        "ipk": "Mencatat pembayaran utang usaha kepada pemasok menggunakan menu Vendor Payment",
        "indikator": "Diberikan data pembayaran utang kepada pemasok, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "PT Karya Mandiri melakukan pembayaran atas utang kepada PT Supplier Utama sebesar Rp12.000.000 menggunakan cek giro. Transaksi pembayaran utang dagang kepada pemasok ini di Accurate Online dicatat melalui menu ...",
        "A": "Menu Penjualan → Penerimaan Pelanggan",
        "B": "Menu Pembelian → Pembayaran Pemasok (Vendor Payment)",
        "C": "Menu Kas & Bank → Penerimaan Kas",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Buku Besar → Saldo Awal",
        "jawaban": "B",
    },
    {
        "no": 14, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Jurnal Umum",
        "cp": "Peserta didik mampu mencatat transaksi menggunakan jurnal umum di Accurate Online",
        "ipk": "Menggunakan menu Jurnal Umum untuk mencatat transaksi yang tidak tersedia di menu khusus",
        "indikator": "Diberikan jenis transaksi yang memerlukan pencatatan manual, peserta didik dapat menentukan menu Jurnal Umum di Accurate Online",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Seorang akuntan perlu mencatat transaksi penyusutan aset tetap dan koreksi saldo akun yang tidak bisa dilakukan melalui menu transaksi khusus di Accurate Online. Untuk mencatat transaksi-transaksi seperti ini secara manual menggunakan jurnal, menu yang digunakan adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Buku Besar → Jurnal Umum (General Journal)",
        "C": "Menu Pembelian → Faktur Pembelian",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Kas & Bank → Transfer Bank",
        "jawaban": "B",
    },
    {
        "no": 15, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Data Master Pemasok",
        "cp": "Peserta didik mampu membuat dan mengelola data master pemasok di Accurate Online",
        "ipk": "Membuat data master pemasok (vendor) di Accurate Online sebelum mencatat transaksi pembelian",
        "indikator": "Diberikan skenario transaksi dengan pemasok baru, peserta didik dapat mengidentifikasi menu yang tepat untuk input data pemasok",
        "tingkat": "Rendah", "nilai": "5",
        "soal": "PT Sejahtera Abadi akan melakukan pembelian perdana dari pemasok baru bernama CV Bahan Baku Prima. Sebelum transaksi pembelian dapat dicatat, data master pemasok tersebut harus dibuat terlebih dahulu. Di Accurate Online, untuk menginput data master pemasok (vendor/supplier) dilakukan melalui ...",
        "A": "Menu Penjualan → Data Pelanggan",
        "B": "Menu Pembelian → Data Pemasok (Vendor List)",
        "C": "Menu Persediaan → Data Barang",
        "D": "Menu Buku Besar → Daftar Akun",
        "E": "Menu Laporan → Laporan Pembelian",
        "jawaban": "B",
    },
    {
        "no": 16, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Metode Pencatatan Persediaan",
        "cp": "Peserta didik mampu menjelaskan metode pencatatan persediaan yang tersedia di Accurate Online",
        "ipk": "Membedakan metode FIFO, LIFO, dan Average dalam pencatatan persediaan",
        "indikator": "Diberikan deskripsi metode pencatatan persediaan, peserta didik dapat mengidentifikasi nama metode yang dimaksud",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Accurate Online mendukung beberapa metode pencatatan persediaan yang dapat dipilih sesuai kebijakan perusahaan. Salah satu metode menyatakan bahwa barang yang pertama kali masuk ke gudang adalah barang yang pertama kali harus dikeluarkan saat terjadi penjualan. Metode pencatatan persediaan tersebut adalah ...",
        "A": "LIFO (Last In First Out)",
        "B": "Average (Rata-rata Bergerak)",
        "C": "FIFO (First In First Out)",
        "D": "Specific Identification (Identifikasi Khusus)",
        "E": "Lower of Cost or Market",
        "jawaban": "C",
    },
    {
        "no": 17, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Penyesuaian Stok (Stock Adjustment)",
        "cp": "Peserta didik mampu melakukan penyesuaian stok barang di Accurate Online berdasarkan hasil stock opname",
        "ipk": "Melakukan stock adjustment di Accurate Online setelah ditemukan selisih stok dari stock opname",
        "indikator": "Diberikan hasil stock opname yang menunjukkan selisih, peserta didik dapat menentukan menu penyesuaian stok",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Setelah dilakukan stock opname, ditemukan bahwa stok fisik barang jenis A hanya 85 unit, sementara catatan di Accurate Online menunjukkan 90 unit. Terdapat selisih kurang sebanyak 5 unit. Untuk menyesuaikan data stok di Accurate Online agar sesuai dengan kondisi fisik, fitur yang digunakan adalah ...",
        "A": "Menu Penjualan → Faktur Penjualan",
        "B": "Menu Pembelian → Purchase Order",
        "C": "Menu Persediaan → Penyesuaian Stok (Stock Adjustment)",
        "D": "Menu Buku Besar → Jurnal Umum",
        "E": "Menu Kas & Bank → Transfer Kas",
        "jawaban": "C",
    },
    {
        "no": 18, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Laporan Buku Besar",
        "cp": "Peserta didik mampu mengakses dan membaca laporan Buku Besar di Accurate Online",
        "ipk": "Mengidentifikasi jenis laporan yang menampilkan rincian mutasi setiap akun",
        "indikator": "Diberikan kebutuhan informasi mutasi akun, peserta didik dapat menentukan jenis laporan yang tepat di Accurate Online",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Manajer keuangan PT Abadi Sejahtera meminta akuntan untuk menyajikan rincian setiap transaksi yang mempengaruhi akun Piutang Usaha selama bulan April 2025, termasuk tanggal, keterangan, dan jumlah setiap transaksi. Laporan yang harus dihasilkan dari Accurate Online adalah ...",
        "A": "Laporan Neraca (Balance Sheet)",
        "B": "Laporan Laba Rugi (Income Statement)",
        "C": "Laporan Buku Besar (General Ledger)",
        "D": "Laporan Arus Kas (Cash Flow Statement)",
        "E": "Laporan Neraca Saldo (Trial Balance)",
        "jawaban": "C",
    },
    {
        "no": 19, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Neraca Saldo (Trial Balance)",
        "cp": "Peserta didik mampu menghasilkan dan menginterpretasikan laporan Neraca Saldo dari Accurate Online",
        "ipk": "Mengakses laporan Neraca Saldo (Trial Balance) di Accurate Online melalui menu Laporan",
        "indikator": "Diberikan kebutuhan laporan neraca saldo, peserta didik dapat menentukan langkah yang tepat untuk mengaksesnya",
        "tingkat": "Sedang", "nilai": "5",
        "soal": "Pada akhir periode akuntansi, akuntan PT Maju Bersama perlu menyusun laporan yang memuat daftar semua akun beserta saldo debit dan kredit masing-masing untuk memverifikasi keseimbangan pencatatan. Di Accurate Online, laporan Neraca Saldo (Trial Balance) dapat diakses melalui ...",
        "A": "Menu Penjualan → Laporan Penjualan",
        "B": "Menu Pembelian → Laporan Pembelian",
        "C": "Menu Laporan → Buku Besar → Neraca Saldo",
        "D": "Menu Persediaan → Laporan Persediaan",
        "E": "Menu Kas & Bank → Laporan Kas",
        "jawaban": "C",
    },
    {
        "no": 20, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Transaksi Kas & Bank - Pengeluaran Biaya",
        "cp": "Peserta didik mampu mencatat transaksi pengeluaran kas untuk biaya operasional di Accurate Online",
        "ipk": "Mencatat pengeluaran kas yang tidak terkait langsung dengan pembelian barang dagangan",
        "indikator": "Diberikan transaksi pengeluaran biaya operasional, peserta didik dapat menentukan menu yang tepat di Accurate Online",
        "tingkat": "Sulit", "nilai": "5",
        "soal": "PT Makmur Jaya membayar biaya sewa gedung kantor sebesar Rp3.000.000 secara tunai. Transaksi ini bukan merupakan pembelian barang dagangan, melainkan pengeluaran biaya operasional perusahaan. Di Accurate Online, pencatatan transaksi pengeluaran biaya operasional seperti ini dilakukan melalui ...",
        "A": "Menu Pembelian → Faktur Pembelian",
        "B": "Menu Penjualan → Faktur Penjualan",
        "C": "Menu Kas & Bank → Pengeluaran Kas/Bank",
        "D": "Menu Persediaan → Penyesuaian Stok",
        "E": "Menu Buku Besar → Saldo Awal",
        "jawaban": "C",
    },
]

# ============================================================
# DATA SOAL ESSAY (10 soal)
# ============================================================
essay_data = [
    {
        "no": 1, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Setup Data Awal Perusahaan di Accurate Online",
        "cp": "Peserta didik mampu melakukan setup data awal perusahaan di Accurate Online secara lengkap dan benar",
        "ipk": "Menjelaskan prosedur pembuatan database dan setup awal perusahaan di Accurate Online",
        "indikator": "Diberikan kasus perusahaan baru yang akan menggunakan Accurate Online, peserta didik dapat menjelaskan langkah-langkah setup awal dengan benar",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Accurate Online merupakan software akuntansi berbasis cloud yang dikembangkan oleh PT Cipta Piranti Sejahtera. Software ini telah banyak digunakan oleh berbagai perusahaan di Indonesia karena kemudahannya dalam mencatat dan mengelola data keuangan. Tidak perlu instalasi khusus karena cukup diakses melalui browser internet.\n\nSebelum mulai menggunakan Accurate Online, pengguna harus melakukan setup atau pengaturan awal perusahaan terlebih dahulu. Proses setup ini meliputi pengisian berbagai informasi dasar perusahaan yang akan menjadi landasan dari semua proses pencatatan akuntansi selanjutnya. Kesalahan dalam proses setup awal dapat berdampak pada keseluruhan sistem pencatatan.\n\nBayangkan kamu adalah seorang staf akuntansi di perusahaan yang baru saja memutuskan untuk menggunakan Accurate Online sebagai software akuntansinya.",
        "pertanyaan": "Jelaskan minimal 5 langkah yang harus dilakukan dalam proses setup data awal perusahaan di Accurate Online beserta penjelasan singkat setiap langkahnya!",
        "jawaban": "Langkah-langkah setup data awal perusahaan di Accurate Online:\n\n1. Login ke Accurate Online - Masuk ke website app.accurate.id menggunakan akun yang telah terdaftar.\n\n2. Buat Database Perusahaan Baru - Klik tombol 'Buat Perusahaan Baru' dan ikuti wizard pembuatan database.\n\n3. Isi Informasi Umum Perusahaan - Masukkan nama perusahaan, alamat lengkap, nomor telepon, email, dan NPWP perusahaan.\n\n4. Tentukan Periode Fiskal - Pilih tahun buku perusahaan, misalnya 1 Januari 2025 s.d. 31 Desember 2025.\n\n5. Pilih Mata Uang - Tentukan mata uang fungsional yang digunakan (misalnya Rupiah/IDR).\n\n6. Pilih Metode Persediaan - Tentukan metode pencatatan persediaan: FIFO (First In First Out) atau Average (Rata-rata Bergerak).\n\n7. Atur Daftar Akun - Sesuaikan Chart of Accounts yang tersedia dengan kebutuhan perusahaan; tambah atau nonaktifkan akun yang diperlukan.\n\n8. Masukkan Saldo Awal - Input saldo awal semua akun berdasarkan neraca perusahaan sebelumnya agar data awal akurat.",
    },
    {
        "no": 2, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Daftar Akun (Chart of Accounts) di Accurate Online",
        "cp": "Peserta didik mampu membuat dan mengelola daftar akun di Accurate Online sesuai kebutuhan perusahaan",
        "ipk": "Menjelaskan pengelompokan akun dalam Accurate Online beserta contoh masing-masing kelompok",
        "indikator": "Diberikan informasi tentang sistem pengkodean akun Accurate Online, peserta didik dapat menjelaskan kelompok akun beserta contoh akun untuk setiap kelompoknya",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Daftar Akun (Chart of Accounts) merupakan komponen fundamental dalam sistem akuntansi berbantuan komputer seperti Accurate Online. Daftar akun berisi seluruh akun yang akan digunakan perusahaan untuk mencatat semua transaksi keuangannya. Setiap akun memiliki kode unik, nama, dan tipe yang berbeda-beda.\n\nAccurate Online telah menyediakan daftar akun standar secara default yang dapat dimodifikasi sesuai kebutuhan spesifik perusahaan. Sistem pengkodean akun menggunakan angka sebagai identifikasi kelompok: angka 1 untuk Aset, 2 untuk Liabilitas, 3 untuk Ekuitas, 4 untuk Pendapatan, dan 5 untuk Beban.\n\nSebagai pelajar Komputer Akuntansi Kelas XI, pemahaman tentang pengelompokan akun sangat penting karena berpengaruh langsung terhadap keakuratan laporan keuangan yang dihasilkan.",
        "pertanyaan": "Jelaskan pengelompokan akun dalam Accurate Online berdasarkan kode akun (kelompok 1 sampai dengan kelompok 5), beserta minimal 3 contoh akun untuk setiap kelompoknya!",
        "jawaban": "Pengelompokan akun dalam Accurate Online:\n\n1. Kelompok 1 - Aset (Harta): Mencerminkan kekayaan yang dimiliki perusahaan.\nContoh akun: Kas (1-10100), Piutang Usaha (1-10300), Persediaan Barang (1-10500), Peralatan Kantor (1-15000), Kendaraan (1-16000).\n\n2. Kelompok 2 - Liabilitas (Kewajiban): Mencerminkan utang/kewajiban perusahaan kepada pihak lain.\nContoh akun: Utang Usaha (2-10100), Utang Bank (2-11000), Utang Gaji (2-10500).\n\n3. Kelompok 3 - Ekuitas (Modal): Mencerminkan hak pemilik atas aset perusahaan.\nContoh akun: Modal Pemilik (3-10000), Laba Ditahan (3-20000), Prive Pemilik (3-30000).\n\n4. Kelompok 4 - Pendapatan: Mencatat seluruh penghasilan yang diperoleh perusahaan.\nContoh akun: Pendapatan Penjualan (4-10000), Pendapatan Jasa (4-20000), Diskon Penjualan (4-30000).\n\n5. Kelompok 5 - Beban: Mencatat seluruh pengeluaran/biaya yang terjadi dalam operasional.\nContoh akun: Harga Pokok Penjualan/HPP (5-10000), Beban Gaji (5-20000), Beban Sewa (5-30000).",
    },
    {
        "no": 3, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Input Saldo Awal Akun di Accurate Online",
        "cp": "Peserta didik mampu menginput saldo awal seluruh akun neraca di Accurate Online untuk perusahaan yang sudah berjalan",
        "ipk": "Menjelaskan prosedur input saldo awal dan prinsip keseimbangan debit-kredit",
        "indikator": "Diberikan kondisi perusahaan yang baru menggunakan Accurate Online, peserta didik dapat menjelaskan prosedur dan prinsip input saldo awal dengan benar",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Ketika sebuah perusahaan yang sudah berjalan mulai menggunakan Accurate Online, salah satu tahapan paling krusial adalah memasukkan saldo awal dari semua akun yang ada. Saldo awal ini berasal dari neraca saldo perusahaan pada akhir periode sebelumnya.\n\nTanpa saldo awal yang benar dan lengkap, laporan keuangan yang dihasilkan oleh Accurate Online tidak akan mencerminkan kondisi keuangan perusahaan yang sesungguhnya. Proses input saldo awal harus dilakukan dengan sangat teliti karena kesalahan sekecil apapun dapat mengakibatkan laporan menjadi tidak akurat.\n\nPrinsip dasar akuntansi menyatakan bahwa dalam setiap pencatatan transaksi, total saldo debit harus selalu sama dengan total saldo kredit. Hal ini berlaku pula dalam input saldo awal di Accurate Online.",
        "pertanyaan": "Jelaskan apa yang dimaksud dengan saldo awal dalam Accurate Online, uraikan prosedur input saldo awal akun secara lengkap, dan jelaskan mengapa total saldo debit harus sama dengan total saldo kredit!",
        "jawaban": "Pengertian Saldo Awal:\nSaldo awal (opening balance) adalah saldo dari setiap akun pada saat pertama kali perusahaan mulai menggunakan Accurate Online, bersumber dari laporan neraca atau neraca saldo akhir periode sebelumnya.\n\nProsedur Input Saldo Awal:\n1. Buka Menu Buku Besar, pilih Saldo Awal\n2. Tentukan tanggal saldo awal (biasanya awal periode penggunaan sistem)\n3. Masukkan saldo masing-masing akun:\n   - Akun Aset: masukkan di kolom Debit\n   - Akun Liabilitas dan Ekuitas: masukkan di kolom Kredit\n4. Periksa total - pastikan Total Debit = Total Kredit\n5. Klik Simpan untuk menyimpan data saldo awal\n\nAlasan Debit Harus = Kredit:\nPrinsip ini merupakan dasar dari sistem akuntansi double-entry bookkeeping (pencatatan ganda). Setiap transaksi keuangan selalu mempengaruhi minimal dua akun dengan jumlah yang sama namun di sisi yang berlawanan (debit dan kredit), sehingga persamaan akuntansi Aset = Liabilitas + Ekuitas selalu terjaga keseimbangannya. Jika total debit tidak sama dengan total kredit, berarti ada kesalahan dalam pencatatan.",
    },
    {
        "no": 4, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pencatatan Transaksi Penjualan Kredit di Accurate Online",
        "cp": "Peserta didik mampu mencatat siklus lengkap transaksi penjualan kredit di Accurate Online",
        "ipk": "Menjelaskan prosedur pencatatan penjualan kredit dari pembuatan faktur hingga pelunasan piutang",
        "indikator": "Diberikan skenario transaksi penjualan kredit, peserta didik dapat menjelaskan prosedur lengkap pencatatannya di Accurate Online beserta jurnal yang terbentuk",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Transaksi penjualan merupakan salah satu transaksi utama yang paling sering dilakukan oleh perusahaan dagang. Di Accurate Online, terdapat alur yang sistematis dalam proses pencatatan penjualan, mulai dari sales order (pesanan penjualan) hingga faktur penjualan (invoice) yang merupakan dokumen resmi tagihan kepada pelanggan.\n\nFaktur penjualan yang dibuat di Accurate Online akan secara otomatis menghasilkan jurnal akuntansi yang sesuai. Sistem juga akan otomatis memperbarui saldo piutang usaha dan mengurangi stok persediaan barang yang terjual.\n\nSiklus penjualan kredit belum selesai sampai pelanggan melunasi pembayarannya. Pelunasan piutang harus dicatat agar saldo piutang usaha menjadi akurat dan laporan keuangan mencerminkan kondisi yang sebenarnya.",
        "pertanyaan": "Jelaskan prosedur lengkap pencatatan transaksi penjualan kredit di Accurate Online, mulai dari pembuatan faktur penjualan hingga pelunasan piutang oleh pelanggan! Sertakan jurnal yang terbentuk pada setiap tahap.",
        "jawaban": "Prosedur Penjualan Kredit di Accurate Online:\n\nTAHAP 1 - Pembuatan Faktur Penjualan:\n1. Buka Menu Penjualan - Faktur Penjualan - Baru\n2. Pilih nama pelanggan dari daftar pelanggan\n3. Isi tanggal transaksi dan nomor faktur\n4. Pilih barang yang dijual, isi kuantitas dan harga jual\n5. Tentukan syarat pembayaran (contoh: Net 30)\n6. Klik Simpan\nJurnal otomatis terbentuk:\n   Debit  : Piutang Usaha\n   Kredit : Penjualan\n   Debit  : HPP (Harga Pokok Penjualan)\n   Kredit : Persediaan Barang\n\nTAHAP 2 - Penerimaan Pembayaran dari Pelanggan:\n1. Buka Menu Penjualan - Penerimaan Pelanggan (Customer Receipt) - Baru\n2. Pilih nama pelanggan\n3. Pilih faktur yang dilunasi dan isi jumlah yang diterima\n4. Pilih akun kas/bank yang menerima pembayaran\n5. Isi tanggal penerimaan dan klik Simpan\nJurnal otomatis terbentuk:\n   Debit  : Kas/Bank\n   Kredit : Piutang Usaha",
    },
    {
        "no": 5, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Penjualan di Accurate Online",
        "cp": "Peserta didik mampu mencatat transaksi retur penjualan di Accurate Online dengan benar",
        "ipk": "Menjelaskan prosedur pencatatan retur penjualan dan jurnal yang terbentuk",
        "indikator": "Diberikan skenario pengembalian barang dari pelanggan, peserta didik dapat menjelaskan prosedur dan jurnal retur penjualan di Accurate Online",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Retur penjualan terjadi ketika pelanggan mengembalikan barang yang telah dibeli karena berbagai alasan, antara lain: barang rusak/cacat saat diterima, barang tidak sesuai spesifikasi yang dipesan, atau kelebihan pengiriman barang.\n\nDi Accurate Online, proses retur penjualan terintegrasi secara langsung dengan faktur penjualan yang menjadi referensinya. Integrasi ini memudahkan pelacakan dan memastikan konsistensi data antara penjualan dan retuurnya.\n\nPencatatan retur penjualan yang tepat dan tepat waktu sangat penting untuk memastikan bahwa saldo piutang usaha, saldo persediaan barang, dan laporan penjualan mencerminkan kondisi yang sebenarnya.",
        "pertanyaan": "Jelaskan prosedur pencatatan retur penjualan di Accurate Online secara lengkap, dan sebutkan jurnal akuntansi yang secara otomatis terbentuk dari transaksi retur penjualan tersebut!",
        "jawaban": "Prosedur Retur Penjualan di Accurate Online:\n\n1. Buka Menu Penjualan - Retur Penjualan - Baru\n2. Pilih nama pelanggan yang melakukan pengembalian barang\n3. Di kolom referensi, pilih nomor faktur penjualan asli yang menjadi dasar retur\n4. Sistem otomatis menampilkan daftar barang dari faktur tersebut\n5. Pilih barang yang diretur dan masukkan kuantitas yang dikembalikan\n6. Isi keterangan/alasan retur (opsional namun disarankan)\n7. Tentukan tanggal retur\n8. Klik Simpan\n\nJurnal yang Terbentuk Otomatis:\nUntuk retur penjualan kredit:\n   Debit  : Retur & Potongan Penjualan (mengurangi pendapatan)\n   Kredit : Piutang Usaha (mengurangi tagihan kepada pelanggan)\n\nUntuk pengembalian stok barang ke gudang:\n   Debit  : Persediaan Barang (barang kembali ke gudang)\n   Kredit : HPP (Harga Pokok Penjualan berkurang)",
    },
    {
        "no": 6, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Pencatatan Transaksi Pembelian Kredit di Accurate Online",
        "cp": "Peserta didik mampu mencatat siklus lengkap transaksi pembelian kredit di Accurate Online",
        "ipk": "Menjelaskan prosedur pencatatan pembelian kredit dari Purchase Order hingga pembayaran utang",
        "indikator": "Diberikan skenario pembelian kredit, peserta didik dapat menjelaskan prosedur dan jurnal lengkap pembelian kredit di Accurate Online",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Transaksi pembelian barang dari pemasok merupakan kegiatan rutin yang dilakukan oleh perusahaan dagang. Di Accurate Online, proses pembelian dapat dimulai dengan pembuatan Purchase Order (PO) sebagai surat pesanan kepada pemasok, dilanjutkan dengan penerimaan barang, dan diakhiri dengan pembuatan faktur pembelian.\n\nPencatatan pembelian yang akurat sangat penting untuk memantau posisi utang kepada pemasok dan memastikan ketersediaan stok barang. Setiap transaksi pembelian di Accurate Online akan otomatis memperbarui saldo utang usaha dan menambah jumlah persediaan barang.\n\nSiklus pembelian kredit belum lengkap sampai utang kepada pemasok dilunasi. Pembayaran utang juga harus dicatat dengan benar agar posisi kas dan saldo utang selalu akurat.",
        "pertanyaan": "Jelaskan prosedur lengkap pencatatan transaksi pembelian kredit di Accurate Online, mulai dari Purchase Order (PO) hingga pembayaran utang kepada pemasok! Sertakan jurnal yang terbentuk pada setiap tahap.",
        "jawaban": "Prosedur Pembelian Kredit di Accurate Online:\n\nTAHAP 1 - Pembuatan Purchase Order (PO):\n1. Buka Menu Pembelian - Purchase Order - Baru\n2. Pilih nama pemasok dan isi tanggal PO\n3. Pilih barang yang dipesan, isi kuantitas dan harga\n4. Klik Simpan\n(PO tidak menghasilkan jurnal, hanya sebagai dokumen pemesanan)\n\nTAHAP 2 - Pembuatan Faktur Pembelian:\n1. Buka Menu Pembelian - Faktur Pembelian - Baru\n2. Pilih pemasok dan referensikan ke PO yang sudah dibuat\n3. Isi nomor faktur dari pemasok dan tanggal jatuh tempo\n4. Klik Simpan\nJurnal otomatis:\n   Debit  : Persediaan Barang\n   Kredit : Utang Usaha\n\nTAHAP 3 - Pembayaran Utang kepada Pemasok:\n1. Buka Menu Pembelian - Pembayaran Pemasok (Vendor Payment) - Baru\n2. Pilih nama pemasok\n3. Pilih faktur yang akan dibayar dan masukkan jumlah pembayaran\n4. Pilih akun kas/bank yang digunakan untuk membayar\n5. Klik Simpan\nJurnal otomatis:\n   Debit  : Utang Usaha\n   Kredit : Kas/Bank",
    },
    {
        "no": 7, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Retur Pembelian di Accurate Online",
        "cp": "Peserta didik mampu mencatat transaksi retur pembelian di Accurate Online dengan benar",
        "ipk": "Menjelaskan prosedur dan jurnal retur pembelian beserta contoh kasusnya",
        "indikator": "Diberikan contoh kasus retur pembelian, peserta didik dapat menjelaskan prosedur pencatatan dan jurnalnya di Accurate Online",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Retur pembelian terjadi ketika perusahaan mengembalikan sebagian atau seluruh barang yang telah dibeli dari pemasok karena: barang cacat/rusak, barang tidak sesuai spesifikasi dalam Purchase Order, atau kelebihan pengiriman dari jumlah yang dipesan.\n\nDi Accurate Online, fitur retur pembelian dirancang untuk terintegrasi langsung dengan faktur pembelian yang menjadi referensinya. Sistem secara otomatis memperbarui saldo utang kepada pemasok dan mengurangi jumlah stok barang yang dikembalikan.\n\nContoh Kasus: CV Mitra Abadi membeli 100 unit barang dari PT Jaya Supplier seharga Rp5.000.000 (Rp50.000/unit) secara kredit. Setelah diperiksa, 10 unit barang dalam kondisi rusak dan dikembalikan kepada PT Jaya Supplier.",
        "pertanyaan": "Berdasarkan contoh kasus di atas, jelaskan prosedur pencatatan retur pembelian di Accurate Online secara lengkap dan sebutkan jurnal yang terbentuk dari transaksi tersebut!",
        "jawaban": "Prosedur Retur Pembelian di Accurate Online (kasus CV Mitra Abadi):\n\n1. Buka Menu Pembelian - Retur Pembelian - Baru\n2. Pilih nama pemasok: PT Jaya Supplier\n3. Di kolom referensi, pilih nomor faktur pembelian asli sebagai dasar retur\n4. Sistem otomatis menampilkan daftar barang yang dibeli\n5. Pilih barang yang diretur, masukkan kuantitas = 10 unit\n6. Sistem otomatis menghitung nilai retur: 10 unit x Rp50.000 = Rp500.000\n7. Isi keterangan alasan retur: 'Barang rusak'\n8. Tentukan tanggal retur pembelian\n9. Klik Simpan\n\nJurnal yang Terbentuk Otomatis:\n   Debit  : Utang Usaha (PT Jaya Supplier)   Rp500.000\n   Kredit : Persediaan Barang               Rp500.000\n\nDampak pencatatan:\n- Utang CV Mitra Abadi berkurang: dari Rp5.000.000 menjadi Rp4.500.000\n- Stok barang berkurang: dari 100 unit menjadi 90 unit",
    },
    {
        "no": 8, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Penyesuaian Stok (Stock Adjustment) di Accurate Online",
        "cp": "Peserta didik mampu melakukan penyesuaian stok berdasarkan hasil stock opname di Accurate Online",
        "ipk": "Menjelaskan prosedur stock opname dan stock adjustment di Accurate Online beserta jurnalnya",
        "indikator": "Diberikan hasil stock opname dengan selisih stok, peserta didik dapat menjelaskan prosedur penyesuaian stok dan jurnal yang terbentuk",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Stock opname adalah kegiatan penghitungan fisik stok barang yang ada di gudang secara langsung, kemudian membandingkan hasilnya dengan data stok yang tercatat di sistem Accurate Online. Stock opname biasanya dilakukan secara berkala, misalnya setiap akhir bulan atau akhir tahun.\n\nPerbedaan antara stok fisik dan stok di sistem dapat terjadi karena: barang hilang, barang rusak yang belum dilaporkan, kesalahan pencatatan, atau barang tercecer. Jika ditemukan selisih, maka harus dilakukan penyesuaian stok (stock adjustment) agar data sistem kembali sesuai kondisi fisik.\n\nContoh Kasus: Hasil stock opname menunjukkan stok fisik Barang X hanya 85 unit, sedangkan catatan Accurate Online menunjukkan 90 unit. Terdapat selisih kurang 5 unit dengan harga pokok Rp20.000/unit.",
        "pertanyaan": "Jelaskan apa yang dimaksud dengan stock opname, uraikan prosedur penyesuaian stok (stock adjustment) di Accurate Online berdasarkan kasus di atas, dan sebutkan jurnal yang terbentuk!",
        "jawaban": "Pengertian Stock Opname:\nStock opname adalah proses penghitungan fisik seluruh persediaan barang di gudang dan membandingkannya dengan catatan di sistem akuntansi untuk memastikan keakuratan data stok.\n\nProsedur Stock Adjustment di Accurate Online (kasus Barang X):\n1. Lakukan penghitungan fisik: Barang X = 85 unit\n2. Bandingkan dengan Accurate Online: tercatat 90 unit\n3. Selisih = 90 - 85 = 5 unit (stok fisik lebih sedikit)\n4. Buka Menu Persediaan - Penyesuaian Stok (Stock Adjustment) - Baru\n5. Pilih barang: Barang X\n6. Masukkan kuantitas aktual: 85 unit\n7. Sistem otomatis menghitung selisih: -5 unit\n8. Isi keterangan: 'Penyesuaian hasil stock opname'\n9. Klik Simpan\n\nJurnal yang Terbentuk:\nNilai selisih = 5 unit x Rp20.000 = Rp100.000\n   Debit  : Beban Kerugian Selisih Persediaan   Rp100.000\n   Kredit : Persediaan Barang                   Rp100.000\n(Mencatat pengurangan stok akibat selisih hasil stock opname)",
    },
    {
        "no": 9, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Neraca Saldo (Trial Balance) di Accurate Online",
        "cp": "Peserta didik mampu menghasilkan dan menginterpretasikan laporan Neraca Saldo dari Accurate Online",
        "ipk": "Menjelaskan pengertian, fungsi, dan cara mengakses Neraca Saldo di Accurate Online",
        "indikator": "Diberikan kebutuhan laporan neraca saldo, peserta didik dapat menjelaskan pengertian, fungsi, dan cara mengaksesnya di Accurate Online",
        "tingkat": "Sedang", "nilai": "10",
        "narasi": "Neraca Saldo (Trial Balance) merupakan salah satu laporan penting yang dihasilkan dalam proses akuntansi. Laporan ini memuat daftar semua akun beserta saldo debit atau kredit masing-masing akun pada suatu periode tertentu.\n\nNeraca saldo menjadi alat verifikasi pertama sebelum laporan keuangan disusun. Jika total saldo debit sama dengan total saldo kredit, maka secara matematis pencatatan transaksi telah seimbang. Namun perlu diingat bahwa neraca saldo yang seimbang tidak menjamin bebas dari semua jenis kesalahan pencatatan.\n\nDi Accurate Online, neraca saldo dihasilkan secara otomatis berdasarkan seluruh transaksi yang telah dicatat dalam sistem dan dapat dicetak atau diekspor ke format Excel untuk keperluan analisis lebih lanjut.",
        "pertanyaan": "Jelaskan pengertian Neraca Saldo (Trial Balance) dalam akuntansi, uraikan minimal 3 fungsinya, dan jelaskan langkah-langkah mengakses laporan Neraca Saldo di Accurate Online!",
        "jawaban": "Pengertian Neraca Saldo:\nNeraca Saldo (Trial Balance) adalah daftar yang memuat semua akun beserta saldo debit atau kredit masing-masing pada akhir suatu periode akuntansi, digunakan untuk memverifikasi keseimbangan antara total debit dan total kredit.\n\nFungsi Neraca Saldo:\n1. Memverifikasi keseimbangan pencatatan: Memastikan total debit = total kredit sebagai bukti kebenaran double-entry bookkeeping\n2. Dasar penyusunan laporan keuangan: Titik awal dalam menyusun Laporan Laba Rugi dan Neraca (Balance Sheet)\n3. Alat deteksi kesalahan: Membantu mengidentifikasi kemungkinan adanya kesalahan posting atau transaksi yang belum dicatat\n4. Gambaran umum keuangan: Memberikan ringkasan posisi saldo seluruh akun secara cepat\n\nLangkah Mengakses Neraca Saldo di Accurate Online:\n1. Klik Menu Laporan (Reports) di panel menu utama\n2. Pilih submenu Buku Besar (General Ledger)\n3. Pilih Neraca Saldo (Trial Balance)\n4. Tentukan periode laporan: isi tanggal awal dan tanggal akhir\n5. Klik tombol Tampilkan / Generate\n6. Laporan Neraca Saldo ditampilkan dan siap dicetak atau diekspor ke Excel",
    },
    {
        "no": 10, "elemen": "Pengelolaan Komputer Akuntansi",
        "materi": "Alur Kerja Accurate Online: Dari Setup hingga Laporan Neraca Saldo",
        "cp": "Peserta didik mampu menjelaskan alur kerja lengkap penggunaan Accurate Online dari setup awal hingga laporan keuangan",
        "ipk": "Mendeskripsikan workflow Accurate Online secara sistematis dan terstruktur",
        "indikator": "Diberikan tugas mendeskripsikan penggunaan Accurate Online, peserta didik dapat menjelaskan alur kerja lengkap dari awal hingga neraca saldo",
        "tingkat": "Sulit", "nilai": "10",
        "narasi": "Accurate Online adalah sistem akuntansi yang dirancang untuk memudahkan proses pencatatan keuangan perusahaan secara terintegrasi. Semua modul dalam Accurate Online saling terhubung, sehingga data yang diinput di satu modul akan otomatis mempengaruhi modul lainnya.\n\nAgar penggunaan Accurate Online berjalan optimal, pengguna harus memahami alur kerja (workflow) yang benar, mulai dari tahap persiapan awal hingga tahap pelaporan. Pemahaman tentang alur kerja ini membantu menghindari kesalahan input dan memastikan semua transaksi tercatat dengan tepat.\n\nSebagai calon tenaga akuntansi profesional, kamu dituntut untuk dapat menguasai dan menjelaskan seluruh alur penggunaan Accurate Online secara sistematis kepada pengguna lain atau kepada manajemen perusahaan.",
        "pertanyaan": "Jelaskan secara sistematis dan lengkap alur kerja (workflow) penggunaan Accurate Online, mulai dari tahap setup awal perusahaan hingga dihasilkannya laporan Neraca Saldo! Sertakan setiap tahapan beserta penjelasan singkatnya.",
        "jawaban": "Alur Kerja (Workflow) Accurate Online:\n\nTAHAP 1 - SETUP AWAL PERUSAHAAN\nBuat database perusahaan baru, isi informasi perusahaan (nama, alamat, NPWP, periode fiskal, mata uang), pilih metode persediaan (FIFO/Average).\n\nTAHAP 2 - PENGATURAN DAFTAR AKUN\nReview dan sesuaikan Chart of Accounts default, tambah/nonaktifkan akun sesuai kebutuhan perusahaan.\n\nTAHAP 3 - INPUT DATA MASTER\n- Buat Daftar Pelanggan (Customer List)\n- Buat Daftar Pemasok (Vendor List)\n- Buat Daftar Barang/Item (Item List)\n\nTAHAP 4 - INPUT SALDO AWAL\nMasukkan saldo awal semua akun, saldo awal persediaan barang, serta saldo awal piutang dan utang dari neraca periode sebelumnya.\n\nTAHAP 5 - PENCATATAN TRANSAKSI HARIAN\n- Penjualan kredit/tunai: Faktur Penjualan\n- Pembelian kredit/tunai: Faktur Pembelian\n- Penerimaan dari pelanggan: Customer Receipt\n- Pembayaran ke pemasok: Vendor Payment\n- Retur penjualan/pembelian bila ada\n- Transaksi lain: Jurnal Umum / Kas & Bank\n\nTAHAP 6 - LAPORAN NERACA SALDO\nAkses Menu Laporan - Buku Besar - Neraca Saldo, tentukan periode laporan, verifikasi total debit = total kredit, cetak/ekspor laporan untuk analisis lebih lanjut.",
    },
]

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def rebuild_cell(cell, text):
    """Completely rebuild cell content with new text, line by line."""
    tc = cell._tc
    # Remove ALL existing paragraphs
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    
    # Add new paragraphs for each line
    lines = text.split('\n') if text else ['']
    for i, line in enumerate(lines):
        from docx.oxml import OxmlElement
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
    """Update header cell: replace placeholder text for Nama Penulis Soal and Mata Pelajaran."""
    for para in cell.paragraphs:
        for run in para.runs:
            t = run.text
            if '..............di isi' in t:
                run.text = t.replace('..............di isi', NAMA_PENULIS)
            elif '.............di isi' in t:
                run.text = t.replace('.............di isi', f'{MAPEL} / Kelas {KELAS}')
            elif '...............' in t and len(t.strip('.')) == 0:
                run.text = NAMA_PENULIS


def get_unique_cells_with_text(table, col, start_row, end_row, search_text):
    """Get unique cells in a column range that contain search_text."""
    seen = set()
    result = []
    for r in range(start_row, end_row):
        try:
            cell = table.rows[r].cells[col]
            tc_id = id(cell._tc)
            if tc_id not in seen and search_text in cell.text:
                seen.add(tc_id)
                result.append((r, cell))
        except Exception:
            pass
    return result


def fill_pg_table(table, q):
    """Fill a PG kartu soal table."""
    # Fix header in row 0
    fix_header_cell(table.rows[0].cells[0])
    
    # Fill Buku Referensi (row 1, unique cell in col 1)
    try:
        ref_cell = table.rows[1].cells[1]
        ref_id = id(ref_cell._tc)
        if ref_id != id(table.rows[0].cells[0]._tc):
            rebuild_cell(ref_cell, f"Buku Acuan/Referensi:\n{REF}")
    except Exception:
        pass
    
    # Fill Buku Referensi row 2
    try:
        r2c1 = table.rows[2].cells[1]
        if id(r2c1._tc) not in {id(table.rows[0].cells[0]._tc)}:
            rebuild_cell(r2c1, f"Buku Acuan/Referensi:\n{REF}")
    except Exception:
        pass
    
    # Fill LEFT COLUMN metadata
    # R2C0: Elemen Kompetensi value
    try:
        cell = table.rows[2].cells[0]
        rebuild_cell(cell, q['elemen'])
    except Exception:
        pass
    
    # R5C0: Capaian Pembelajaran value (second CP row - value)
    try:
        c4 = table.rows[4].cells[0]
        c5 = table.rows[5].cells[0]
        # R4C0 = "Capaian Pembelajaran:" label, R5C0 = value
        if id(c4._tc) != id(c5._tc):
            rebuild_cell(c5, q['cp'])
        else:
            # If merged, append to existing
            pass
    except Exception:
        pass
    
    # R6C0: Materi
    try:
        cell = table.rows[6].cells[0]
        rebuild_cell(cell, f"Materi: {q['materi']}")
    except Exception:
        pass
    
    # R7C0: Kelas/Semester
    try:
        c6 = table.rows[6].cells[0]
        c7 = table.rows[7].cells[0]
        if id(c6._tc) != id(c7._tc):
            rebuild_cell(c7, f"Kelas: {KELAS} / Semester: Genap")
    except Exception:
        pass
    
    # R10C0 (IPK value - may be duplicate of R9C1)
    try:
        c9 = table.rows[9].cells[0]
        c10 = table.rows[10].cells[0]
        if id(c9._tc) != id(c10._tc):
            rebuild_cell(c10, q['ipk'])
    except Exception:
        pass
    
    # R11C0 (Indikator soal label cell - spans R11-13)
    try:
        cell = table.rows[11].cells[0]
        tingkat = q.get('tingkat', 'Sedang')
        content = (f"Indikator Soal / Indikator Asessmen:\n"
                   f"{q['indikator']}\n\n"
                   f"Tingkat Kesukaran:\n"
                   f"{'[v]' if tingkat=='Rendah' else '[ ]'} Rendah\n"
                   f"{'[v]' if tingkat=='Sedang' else '[ ]'} Sedang\n"
                   f"{'[v]' if tingkat=='Sulit' else '[ ]'} Sulit")
        rebuild_cell(cell, content)
    except Exception:
        pass
    
    # R9C0 (IPK label - keep but update if not merged with other)
    try:
        cell = table.rows[9].cells[0]
        rebuild_cell(cell, f"Indikator Pencapaian Kompetensi:\n{q['ipk']}")
    except Exception:
        pass
    
    # RIGHT SIDE - No. Soal (R4C2 which spans R4-R5 maybe)
    try:
        c4c2 = table.rows[4].cells[2]
        # Fill No. Soal cell - this spans the label "No. Soal" and the number
        rebuild_cell(c4c2, f"No. Soal\n{q['no']}")
    except Exception:
        pass
    
    # R8C2: Nilai
    try:
        c8c2 = table.rows[8].cells[2]
        # Check not a duplicate
        seen_ids = {id(table.rows[r].cells[2]._tc) for r in range(4, 8)}
        if id(c8c2._tc) not in seen_ids:
            rebuild_cell(c8c2, f"{q.get('nilai', '5')}")
    except Exception:
        pass
    
    # R11C2, R12C2, R13C2: Kunci Jawaban
    # Fill R11C2 with answer, clear R12C2 and R13C2
    kunci_cells = []
    seen_kc = set()
    for r in [11, 12, 13]:
        try:
            cell = table.rows[r].cells[2]
            cid = id(cell._tc)
            if cid not in seen_kc:
                seen_kc.add(cid)
                kunci_cells.append((r, cell))
        except Exception:
            pass
    
    if kunci_cells:
        _, first_kunci = kunci_cells[0]
        rebuild_cell(first_kunci, f"KUNCI\nJAWABAN\n\n{q['jawaban']}")
        for _, kc in kunci_cells[1:]:
            rebuild_cell(kc, "")
    
    # QUESTION AREA (col 3): Fill first unique cell, clear others
    # Build the question text
    soal_text = (
        f"{q['soal']}\n\n"
        f"A. {q['A']}\n\n"
        f"B. {q['B']}\n\n"
        f"C. {q['C']}\n\n"
        f"D. {q['D']}\n\n"
        f"E. {q['E']}"
    )
    
    seen_q = set()
    first_q_done = False
    for r in range(3, 14):
        try:
            cell = table.rows[r].cells[3]
            cid = id(cell._tc)
            if cid not in seen_q and 'Ketika Soal' in cell.text:
                seen_q.add(cid)
                if not first_q_done:
                    rebuild_cell(cell, soal_text)
                    first_q_done = True
                else:
                    rebuild_cell(cell, "")
        except Exception:
            pass


def fill_essay_table(table, q):
    """Fill an essay kartu soal table."""
    # Fix header
    fix_header_cell(table.rows[0].cells[0])
    
    # Fill Buku Referensi
    try:
        rebuild_cell(table.rows[1].cells[1], f"Buku Acuan/Referensi:\n{REF}")
    except Exception:
        pass
    try:
        r2c1 = table.rows[2].cells[1]
        if id(r2c1._tc) != id(table.rows[1].cells[1]._tc):
            rebuild_cell(r2c1, f"Buku Acuan/Referensi:\n{REF}")
    except Exception:
        pass
    
    # Fill metadata
    try:
        rebuild_cell(table.rows[2].cells[0], q['elemen'])
    except Exception:
        pass
    
    try:
        c5 = table.rows[5].cells[0]
        rebuild_cell(c5, q['cp'])
    except Exception:
        pass
    
    try:
        rebuild_cell(table.rows[6].cells[0], f"Materi: {q['materi']}")
    except Exception:
        pass
    
    try:
        c6 = table.rows[6].cells[0]
        c7 = table.rows[7].cells[0]
        if id(c6._tc) != id(c7._tc):
            rebuild_cell(c7, f"Kelas: {KELAS} / Semester: Genap")
    except Exception:
        pass
    
    try:
        c9 = table.rows[9].cells[0]
        rebuild_cell(c9, f"Indikator Pencapaian Kompetensi:\n{q['ipk']}")
    except Exception:
        pass
    
    try:
        c10 = table.rows[10].cells[0]
        c9 = table.rows[9].cells[0]
        if id(c10._tc) != id(c9._tc):
            rebuild_cell(c10, q['ipk'])
    except Exception:
        pass
    
    try:
        cell = table.rows[11].cells[0]
        tingkat = q.get('tingkat', 'Sedang')
        rebuild_cell(cell, 
            f"Indikator Soal / Indikator Asessmen:\n{q['indikator']}\n\n"
            f"Tingkat Kesukaran:\n"
            f"{'[v]' if tingkat=='Rendah' else '[ ]'} Rendah\n"
            f"{'[v]' if tingkat=='Sedang' else '[ ]'} Sedang\n"
            f"{'[v]' if tingkat=='Sulit' else '[ ]'} Sulit")
    except Exception:
        pass
    
    # No. Soal
    try:
        rebuild_cell(table.rows[4].cells[2], f"No. Soal\n{q['no']}")
    except Exception:
        pass
    
    # Nilai
    try:
        c8c2 = table.rows[8].cells[2]
        seen_ids = {id(table.rows[r].cells[2]._tc) for r in range(4, 8)}
        if id(c8c2._tc) not in seen_ids:
            rebuild_cell(c8c2, f"{q.get('nilai', '10')}\n(poin)")
    except Exception:
        pass
    
    # Kunci Jawaban label
    seen_kc = set()
    for r in [11, 12, 13]:
        try:
            cell = table.rows[r].cells[2]
            cid = id(cell._tc)
            if cid not in seen_kc:
                seen_kc.add(cid)
                rebuild_cell(cell, "KUNCI\nJAWABAN")
        except Exception:
            pass
    
    # Question + Answer in col 3
    soal_text = (
        f"{q['narasi']}\n\n"
        f"PERTANYAAN:\n{q['pertanyaan']}\n\n"
        f"---KUNCI JAWABAN---\n\n{q['jawaban']}"
    )
    
    seen_q = set()
    first_q_done = False
    for r in range(3, 14):
        try:
            cell = table.rows[r].cells[3]
            cid = id(cell._tc)
            if cid not in seen_q and ('Ketika Soal' in cell.text or 'Ketik Jawaban' in cell.text):
                seen_q.add(cid)
                if not first_q_done:
                    rebuild_cell(cell, soal_text)
                    first_q_done = True
                else:
                    rebuild_cell(cell, "")
        except Exception:
            pass


def fill_kisi_kisi(table):
    pg_groups = [
        (1, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu menjelaskan karakteristik Accurate Online, melakukan setup perusahaan, dan mengelola daftar akun",
         "Pengenalan dan Setup Accurate Online",
         "Mengidentifikasi fitur dan menu Accurate Online untuk pengelolaan data akuntansi",
         "5", "1 - 5", "Pilihan Ganda", "Rendah - Sedang"),
        (2, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu mencatat faktur penjualan, retur penjualan, dan penerimaan dari pelanggan",
         "Transaksi Penjualan, Retur Penjualan, dan Penerimaan",
         "Mencatat siklus transaksi penjualan kredit hingga pelunasan piutang",
         "5", "6 - 10", "Pilihan Ganda", "Sedang"),
        (3, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu mencatat faktur pembelian, retur pembelian, dan pembayaran kepada pemasok",
         "Transaksi Pembelian, Retur Pembelian, dan Pembayaran",
         "Mencatat siklus transaksi pembelian kredit hingga pelunasan utang",
         "5", "11 - 15", "Pilihan Ganda", "Sedang"),
        (4, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu melakukan penyesuaian stok, mencatat jurnal umum, dan menghasilkan laporan neraca saldo",
         "Persediaan, Jurnal Umum, dan Laporan Neraca Saldo",
         "Menggunakan Stock Adjustment, General Journal, dan mengakses Trial Balance",
         "5", "16 - 20", "Pilihan Ganda", "Sedang - Sulit"),
    ]
    essay_groups = [
        (1, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu menjelaskan prosedur setup awal dan pengelolaan daftar akun di Accurate Online",
         "Setup Data Awal, Daftar Akun, dan Saldo Awal",
         "Mendeskripsikan langkah setup perusahaan, kelompok akun, dan saldo awal",
         "3", "1 - 3", "Uraian/Essay", "Sedang"),
        (2, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu menjelaskan siklus penjualan kredit, retur penjualan, dan pembelian kredit",
         "Penjualan Kredit, Retur Penjualan, Pembelian Kredit",
         "Menjelaskan prosedur dan jurnal transaksi penjualan dan pembelian",
         "3", "4 - 6", "Uraian/Essay", "Sedang"),
        (3, "Pengelolaan Komputer Akuntansi",
         "Peserta didik mampu menjelaskan retur pembelian, penyesuaian stok, dan neraca saldo",
         "Retur Pembelian, Stock Adjustment, dan Neraca Saldo",
         "Menjelaskan prosedur retur pembelian, stock adjustment, dan laporan neraca saldo",
         "4", "7 - 10", "Uraian/Essay", "Sulit"),
    ]
    
    for no, elemen, cp, materi, tp, jml, nomor, bentuk, tingkat in pg_groups:
        r = no  # rows 1-4
        row = table.rows[r]
        vals = [str(no), elemen, cp, materi, tp, jml, nomor, bentuk, tingkat]
        for c, v in enumerate(vals):
            try:
                rebuild_cell(row.cells[c], v)
            except Exception:
                pass
    
    for no, elemen, cp, materi, tp, jml, nomor, bentuk, tingkat in essay_groups:
        r = no + 5  # rows 6-8
        row = table.rows[r]
        vals = [str(no), elemen, cp, materi, tp, jml, nomor, bentuk, tingkat]
        for c, v in enumerate(vals):
            try:
                rebuild_cell(row.cells[c], v)
            except Exception:
                pass
    
    # Total row
    try:
        row = table.rows[11]
        rebuild_cell(row.cells[0], "TOTAL")
        rebuild_cell(row.cells[5], "30")
        rebuild_cell(row.cells[6], "PG: 1-20\nEssay: 1-10")
        rebuild_cell(row.cells[7], "PG & Essay")
        rebuild_cell(row.cells[8], "Variatif")
    except Exception:
        pass


def fill_table0(table):
    cell = table.rows[0].cells[0]
    for para in cell.paragraphs:
        for run in para.runs:
            if '...............................' in run.text:
                run.text = run.text.replace('...............................', NAMA_PENULIS)
        full = para.text
        if 'Mata Pelajaran' in full:
            for run in para.runs:
                if '......' in run.text and run.text.strip('.') == '':
                    run.text = MAPEL
        elif 'Kelas' in full and '......' in full:
            for run in para.runs:
                if '......' in run.text and run.text.strip('.') == '':
                    run.text = KELAS


# ============================================================
# MAIN
# ============================================================
def main():
    doc = Document('/root/.claude/uploads/19d916f2-492a-4a0c-bfba-2d98a5361dd0/faf676bf-KARTU_SOAL_UasGENAP_2526okeaslikosongan.docx')
    
    # Add 7 more PG tables (clone from table 2)
    print("Adding extra PG tables...")
    ref_elem = doc.tables[2]._element
    last_pg = doc.tables[14]._element
    parent = last_pg.getparent()
    children = list(parent)
    idx14 = children.index(last_pg)
    
    for i in range(7):
        new_t = deepcopy(ref_elem)
        parent.insert(idx14 + 1 + i, new_t)
    
    print(f"Tables: {len(doc.tables)}")
    
    print("Filling table 0 (general header)...")
    fill_table0(doc.tables[0])
    
    print("Filling table 1 (kisi-kisi)...")
    fill_kisi_kisi(doc.tables[1])
    
    print("Filling PG tables (2-21)...")
    for i, q in enumerate(pg_data):
        tbl_idx = i + 2
        print(f"  Soal PG {q['no']} -> Table {tbl_idx}")
        fill_pg_table(doc.tables[tbl_idx], q)
    
    print("Filling Essay tables (22-31)...")
    for i, q in enumerate(essay_data):
        tbl_idx = 22 + i
        print(f"  Soal Essay {q['no']} -> Table {tbl_idx}")
        fill_essay_table(doc.tables[tbl_idx], q)
    
    output = '/home/user/daw/KARTU_SOAL_UAS_GENAP_KomputerAkuntansi_XI.docx'
    doc.save(output)
    print(f"\nSaved: {output}")

main()
